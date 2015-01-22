using System;
using System.Collections.Generic;
using Ivi.Visa.Interop;
using IronPython.Hosting;
using Microsoft.Scripting;
using Microsoft.Scripting.Hosting;
using System.Text.RegularExpressions;


namespace ArbScripter {
    public static class Strings {
        public const string usage =
              "Usage: ArbScripter <parameter options>\n"
            + "where <parameter options> is 0 or more of:\n"
            + " -i <instrument address or alias>\n"
            + " -e <python code to be evaluated>\n"
            + " -f <filename of script to be interpreted>\n"
            + "\n"
            + "Script files are line oriented. Use '\' at line end for continuation.\n"
            + "\n"
            + "The leading character determines action taken for a logical line:\n"
            + " '!' => send following text immediately as an instrument command. [a]\n"
            + " '&' => send following text as a buffered instrument command. [a]\n"
            + " '*' => send following text with length-prefixed parameter block. [a,b]\n"
            + " '?' => send following text immediately as an instrument query. [c]\n"
            + " '=' => evaluate following text or block as IronPython code. [d]\n"
            + " '@' => send following text appended by vector as binary block. [a,e]\n"
            + " '>' => output following text with brace-enclosed variables to console.\n"
            + " '<' => treat following text as filename and interpret file's lines.\n"
            + " '%' => following text is comment, no action to be taken.\n"
            + " '.' => terminate script execution. (good for debugging)\n"
            + "\n"
            + "[a. Within the text, brace-enclosed identifiers will be replaced,\n"
            + " (along with the braces), with the string representation of an object\n"
            + " referenced in the Python global namespace by the given identifier.\n"
            + " It is an abort-inducing error for the identifier to be undefined. ]\n"
            + "\n"
            + "[b. Parameter block after '#' will be prefixed by its length coded in\n"
            + " Keysight's strange <digit><length> format, where <digit> is a single\n"
            + " digit specifying the number of digits in the base-10 ASCII <length>,\n"
            + " which specifies the number of characters in remaining command text. ]\n"
            + "\n"
            + "[c. Text returned by instrument sent to stdout without other effect. ]\n"
            + "\n"
            + "[d. Text on same line is treated as a single-line code block. When the\n"
            + " remaining line is blank, following lines, up to but not including a\n"
            + " lone '_' character, are collected and evaluated as the code block. ]\n"
            + "\n"
            + "[e. The vector name may be specified with a brace-enclosed identifier\n"
            + " at the text's end, or the default name, 'samples', will be used. ]\n"
            + "\n"
            + "Python code block executions share the same global namespace, and may\n"
            + "access an instrument access object named 'arb' having these members:\n"
            + " void Command(string cmd, bool flush = true)\n"
            + " string Query(string cmd)\n"
            + " bool SendFloats(string lead, List<float> vf)\n"
            + " void SendBlock(string lead, float[] z)\n"
            + " bool SendWithParamLength(string text)\n"
            + " const string noErrString\n"
            ;
        public static Regex varName = new Regex("\\{(\\w+)\\}");
        public static Regex tailVarName = new Regex("\\{(\\w+)\\}$");
    }

    public class FloatList : List<Single> {
        public static FloatList New(int capacity = 0){
            return (capacity > 0)? new FloatList(capacity) : new FloatList();
        }

        FloatList() {
        }
        FloatList(int capacity) : base(capacity) {
        }
        public override string ToString() {
            var sb = new System.Text.StringBuilder(6 + Count * 9 + (Count/8) + 1);
            sb.Append("[\n");
            for (int line = 0; line * 8 < Count; ++line) {
                for (int i = line*8; i < Count && i < line*8 + 8; ++i) {
                    sb.AppendFormat(" {0:+0.00000; 0.00000;-0.00000}", this[i]);
                }
                sb.Append("\n");
            }
            sb.Append("]\n");
            return sb.ToString();
        }
    }
    public class Instrument : IDisposable {
        IResourceManager IRM;
        IVisaSession IVS;
        IFormattedIO488 oFio;
        bool pendingCmd;
        bool faking;

        public Instrument() {
            IRM = new Ivi.Visa.Interop.ResourceManager();
            oFio = new FormattedIO488();
            pendingCmd = false;
            faking = false;
        }
        public bool Open(string instAddress, int wtime = 5000) {
            if (instAddress.ToLower() == "fake") {
                faking = true;
                return true;
            }
            try {
                IVS = IRM.Open(instAddress, AccessMode.NO_LOCK);
                if (IVS != null) {
                    oFio.IO = (IMessage)IVS;
                    oFio.IO.Timeout = wtime;
                    return true;
                }
            }
            catch (Exception) {
            }
            return false;
        }
        public void Close() {
            if (!faking)
            {
                IVS.Close();
                IVS = null;
                oFio.IO = null;
            }
            else
                faking = false;
            pendingCmd = false;
        }
        public bool IsOpen {
            get { return IVS != null || faking; }
        }
        protected virtual void Dispose(bool disposing) {
            if (disposing && IsOpen)
                Close();
        }
        public void Dispose() {
            Dispose(true);
        }
        public void Command(string cmd, bool flush = true) {
            if (pendingCmd && cmd.Length > 0 && !faking) {
                oFio.WriteString("\n", false);
            }
            if (!faking)
                oFio.WriteString(cmd, flush);
            pendingCmd = !flush;
        }
        public string Query(string cmd, string sayIfFake = noErrString) {
            Command(cmd, true);
            if (!faking)
                return oFio.ReadString();
            else
                return sayIfFake;
        }
        public bool SendFloats(string lead, List<float> vf) {
            if (IsOpen) {
                SendBlock(lead, vf.ToArray());
                return true;
            }
            else
                return false;
        }
        public void SendBlock(string lead, float[] z) {
            if (faking)
                return;
            if (pendingCmd) {
                oFio.WriteString("\n", false);
            }
            if ((z == null) || z.Length == 0)
                return;
            oFio.WriteIEEEBlock(lead + " ", z, true);
            // Wait for the operation to complete before moving on.
            oFio.WriteString("*WAI", true);
            pendingCmd = false;
        }
        public bool SendWithParamLength(string text) {
            if (IsOpen) {
                int isharp = text.IndexOf('#');
                if (isharp >= 0) {
                    string lead = text.Substring(0, isharp + 1);
                    string tail = text.Substring(isharp + 1);
                    int ntc = tail.Length;
                    string nPrefix = ntc.ToString();
                    string npp = (nPrefix.Length).ToString();
                    string cmd = lead + npp + nPrefix + tail;
                    Command(cmd, true);
                    return true;
                }
            }
            return false;
        }
        public const string noErrString = "+0,\"No error\"\n";
    }
#if USE_FUNKY_VARS
    public class Vars {
        public double Pi { get { return Math.PI; } }
        public Int32 TimeZeroSampleIndex { get; set; }
        public Int32 GeneratorSampleFrequency { get; set; }
        public double t(Int32 ns) {
            return ((double)(ns - TimeZeroSampleIndex)) / GeneratorSampleFrequency;
        }
        public double Amplitude { get; set; }
        public double F { get; set; }
        public double Fs { get; set; }
        public double Fo { get; set; }
        public double Fd { get; set; }
        public double W {
            get { return 2.0 * Pi * F; }
            set { F = value / (2.0 * Pi); }
        }
        public double Ws {
            get { return 2.0 * Pi * Fs; }
            set { Fs = value / (2.0 * Pi); }
        }
        public double Wo {
            get { return 2.0 * Pi * Fo; }
            set { Fo = value / (2.0 * Pi); }
        }
        public double Wd {
            get { return 2.0 * Pi * Fd; }
            set { Fd = value / (2.0 * Pi); }
        }
        public FloatList samples  = new FloatList();
    }
#endif
    class Program : IDisposable {
        private ScriptEngine m_engine = Python.CreateEngine();
        private ScriptScope m_scope;
#if USE_FUNKY_VARS
        Vars vars;
#endif
        Instrument arb;
        private int compileCount = 0;

        string varSubstitute(string sIn) {
            var m = Strings.varName.Matches(sIn);
            int locAdjust = 0;
            foreach (Match vn in m) {
                string vname = ((System.Text.RegularExpressions.Match)(vn.Captures[0])).Groups[1].Value;
                int vloc = vn.Index + locAdjust;
                dynamic vval = null;
                if (m_scope.TryGetVariable(vname, out vval)) {
                    string vput = vval.ToString();
                    sIn = sIn.Remove(vloc, vn.Value.Length).Insert(vloc, vput);
                    locAdjust += vput.Length - vn.Value.Length;
                }
            }
            return sIn;
        }
        float[] tailArrayFetch(string tail, out string strippedTail, string ifAbsent = null) {
            string an = ifAbsent;
            Match m = Strings.tailVarName.Match(tail);
            if (m == null) {
                strippedTail = tail;
                if (an == null)
                    return null;
            }
            else {
                strippedTail = tail.Remove(m.Index, m.Value.Length);
                an = m.Groups[1].Value;
            }
            dynamic vval;
            if (m_scope.TryGetVariable(an, out vval)) {
                try {
                    List<float> fa = vval as List<float>;
                    return fa.ToArray(); // Might be null.
                }
                catch (Exception) {
                    // Who cares, it was wrong!
                }
            }
            return null;
        }

        bool Interpret(System.IO.TextReader lineSource, string sourceTag) {
            bool done = false;
            bool rv = true;
            int lineCount = 0;
            while (!done) {
                var line = lineSource.ReadLine();
                ++lineCount;
                while (line != null && line.Length > 0 && line[line.Length - 1] == '\\') {
                    line = line.Remove(line.Length - 1);
                    var addendum = lineSource.ReadLine();
                    ++lineCount;
                    if (addendum != null)
                        line = line + addendum;
                }
                done = line == null || (line.Length == 1 && line[0] == '.');
                if (!done && line.Length > 0) {
                    char c = line[0];
                    string tail = (line.Substring(1)).Trim();
                    switch (c) {
                        case '.':
                            done = true;
                            break;
                        case '*':
                            arb.SendWithParamLength(varSubstitute(tail));
                            break;
                        case '!': case '&':
                            if (arb.IsOpen)
                                arb.Command(varSubstitute(tail), c == '!');
                            break;
                        case '@':
                            if (arb.IsOpen) {
#if USE_FUNKY_VARS
                                arb.SendBlock(tail, vars.samples.ToArray());
#else
                                string strippedTail;
                                float[] fa = tailArrayFetch(tail, out strippedTail, "samples");
                                if (fa != null)
                                    arb.SendBlock(varSubstitute(strippedTail), fa);
                                else
                                    Console.WriteLine("Error: Lookup failed in '{0}'", tail);
#endif
                            }
                            break;
                        case '?':
                            if (arb.IsOpen) {
                                string reply = arb.Query(tail);
                                Console.WriteLine("Reply: {0}", reply);
                            }
                            break;
                        case '>': {
                                Console.WriteLine("{0}", varSubstitute(tail));
                            }
                            break;
                        case '=':
                            {
                                var code = tail;
                                var lineSay = lineCount;
                                if (code.Length == 0)
                                {
                                    var lines = new List<string>();
                                    string lin;
                                    do
                                    {
                                        lin = lineSource.ReadLine();
                                        ++lineCount;
                                        if (lin != null)
                                        {
                                            if (lin.Length > 0 && lin[0] == '_')
                                                break;
                                            lines.Add(lin);
                                        }
                                    } while (lin != null);
                                    if (lines.Count > 0 && lin != null)
                                        code = String.Join("\n", lines);
                                }
                                if (code.Length > 0)
                                {
                                    try
                                    {
                                        object o = m_engine.Execute(code, m_scope);
                                    }
                                    catch (Microsoft.Scripting.SyntaxErrorException see)
                                    {
                                        Console.WriteLine("Error: Bad syntax in line {0}, column {1}", see.Line, see.Column);
                                        Console.WriteLine("Code at line {0} of {1}, block #{2}.", lineSay, sourceTag, compileCount);
                                        Console.WriteLine("(Could not compile |{0}|.)", see.SourceCode);
                                    }
                                    catch (IronPython.Runtime.Exceptions.SystemExitException see)
                                    {
                                        Console.WriteLine("Error exit: {0}", see.Message);
                                        done = true;
                                        rv = false;
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine(ex.Message);
                                    }
                                }
                            }
                            break;
                        case '%':
                            break;
                        case '<':
                            if (tail.Length > 0) {
                                System.IO.TextReader fin = null;
                                try {
                                    fin = System.IO.File.OpenText(tail);
                                }
                                catch (System.IO.FileNotFoundException e) {
                                    rv = false;
                                    done = true;
                                    Console.WriteLine("ERROR: {0}", e.Message);
                                }

                                if (fin != null) {
                                    using (fin) {
                                        if (!Interpret(fin, tail))
                                        {
                                            rv = false;
                                            done = true;
                                        }
                                    }
                                    fin.Close();
                                }
                            }
                            break;
                        default:
                            continue;
                    }
                }
            }
            return rv;
        }

        // This delegate acts like a FloatList class contructor in the IronPython execution context.
        // There, it allows the expression 'FloatList()' or 'FloatList(capacity)' to return an object
        // acceptable as a List<Single> and whose string representation shows the contained values. 
        public delegate FloatList NewFloatList(int cap = 0);
        Program() {
            m_scope = m_engine.CreateScope();
            arb = new Instrument();
            m_scope.SetVariable("arb", arb);
            m_scope.SetVariable("FloatList", new NewFloatList(FloatList.New));
#if USE_FUNKY_VARS
            vars = new Vars();
            m_scope.SetVariable("vars", vars);
#else
            m_scope.SetVariable("samples", FloatList.New());
#endif
            string xpath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string xdir = System.IO.Directory.GetParent(xpath).FullName;
            var spath = new List<string>();
            spath.Add(xdir);
            string ipaths = System.Environment.GetEnvironmentVariable("IRONPYTHONPATH");
            if (ipaths != null && ipaths.Length > 0)
            {
                string[] ipa = ipaths.Split(';');
                foreach (string d in ipa)
                {
                    if (d.Length > 0)
                        spath.Add(d);
                }
            }
            m_engine.SetSearchPaths(spath.ToArray());
        }
        bool Run(string[] args) {
            bool rv = true;
            bool done = false;
            for (int i = 0; !done && i < args.Length; ++i) {
                if (!(i + 1 < args.Length)) {
                    Console.WriteLine("Option {0} has no parameter.", args[i]);
                    break;
                }
                string param = args[++i];
                switch (args[i-1]) {
                    case "-e":
                        try {
                            object o = m_engine.Execute(param, m_scope);
                        }
                        catch (IronPython.Runtime.Exceptions.SystemExitException see) {
                            Console.WriteLine("Error exit: {0}", see.Message);
                            done = true;
                            rv = false;
                        }
                        catch (Microsoft.Scripting.SyntaxErrorException see) {
                            Console.WriteLine("Error: Bad syntax in argument {0}, column {1}", i, see.Column);
                            Console.WriteLine("(Could not compile |{0}|.)", see.SourceCode);
                        }
                        catch (Exception ex) {
                            Console.WriteLine(ex.Message);
                        }
                        break;
                    case "-f":
                        if (param == "-")
                            rv = rv && Interpret(Console.In, "<StdInput>");
                        else {
                            System.IO.TextReader fin = null;
                            try {
                                fin = System.IO.File.OpenText(param);
                            }
                            catch (System.IO.FileNotFoundException e) {
                                rv = false;
                                done = true;
                                Console.WriteLine("ERROR: {0}", e.Message);
                            }
                            if (fin != null) {
                                using (fin) {
                                    if (!Interpret(fin, param)) {
                                        rv = false;
                                        done = true;
                                    }
                                }
                                fin.Close();
                            }
                        }
                        break;
                    case "-i":
                        if (arb.IsOpen)
                            arb.Close();
                        if (!arb.Open(param)) {
                            Console.Error.WriteLine("Cannot open {0}", param);
                            return false;
                        }
                        break;
                    case "-?":
                    case "-h":
                    case "--help":
                        Console.WriteLine(Strings.usage);
                        done = true;
                        break;
                    default:
                        break;
                }
            }
            if (rv && !arb.IsOpen) {
                Console.Error.WriteLine("Instrument not opened.");
                return false;
            }
            return rv;
        }
        static int Main(string[] args) {
            if (args.Length == 1)
            {
                var reHelp = new System.Text.RegularExpressions.Regex("^((--?)|/)(([hH](elp)?)|\\?)$");
                if (reHelp.IsMatch(args[0]))
                {
                    Console.Write(Strings.usage);
                    return 0;
                }
            }
            using (var p = new Program())
                return p.Run(args) ? 0 : 2;
        }
        protected virtual void Dispose(bool disposing) {
            if (disposing && arb.IsOpen)
                arb.Close();
        }
        public void Dispose() {
            Dispose(true);
        }

        //public static void BinaryArb()
        //{
        //    var inst = new Instrument();
        //    const int NUM_DATA_POINTS = 100000;
        //    float[] z = new float[NUM_DATA_POINTS];
        //    //Create simple ramp waveform
        //    for (int i = 0; i < NUM_DATA_POINTS; i++)
        //        z[i] = (i - 1) / (float)NUM_DATA_POINTS;

        //    string instAddress = "USB0::2391::19207::MY53400207::0::INSTR";
        //    instAddress = "USB_Arb33612A";
        //    if (inst.Open(instAddress)) {
        //        try {
        //            string reply;
        //            reply = inst.Query("*IDN?");
        //            Console.WriteLine("Instrument Identity String: " + reply);
        //            //Clear and reset instrument
        //            inst.Query("*CLS;*RST;*OPC?");
        //            //Clear volatile memory
        //            inst.Command("SOURce1:DATA:VOLatile:CLEar", true);
        //            // swap the endian format
        //            inst.Command("FORM:BORD NORM", true);
        //            //Downloading 
        //            Console.WriteLine("Downloading Waveform...");
        //            inst.SendBlock("SOURce1:DATA:ARBitrary testarb,", z);
        //            Console.WriteLine("Download Complete", true);
        //            //Set desired configuration
        //            inst.Command("SOURce1:FUNCtion:ARBitrary testarb", false); // set current arb waveform to defined arb pulse
        //            inst.Command("SOURce1:FUNCtion ARB", false); // turn on arb function
        //            inst.Command("SOURCE1:FUNCtion:ARB:SRATe 400000", false); // set sample rate
        //            inst.Command("SOURCE1:VOLT 2", false); // set max waveform amplitude to 2 Vpp
        //            inst.Command("SOURCE1:VOLT:OFFSET 0", false); // set offset to 0 V
        //            inst.Command("OUTPUT1:LOAD 50", false); // set output load to 50 ohms
        //            //Enable Output
        //            inst.Command("OUTPUT1 ON", true); // turn on channel 1 output
        //            //Read Error/s
        //            reply = inst.Query("SYSTEM:ERROR?");
        //            if (reply == Instrument.noErrString) {
        //                Console.WriteLine("Output set without any error\n");
        //            }
        //            else {
        //                Console.WriteLine("Error reported: " + reply);
        //            }
        //        }
        //        finally {
        //            inst.Close();
        //        }
        //    }
        //}        
    }
}
