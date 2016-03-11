using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExposerV2.ConsoleTest
{
    using System.Diagnostics;
    using System.Reflection.Emit;
    using System.Threading;

    using FluentIL;
    using FluentIL.Infos;

    using Jurassic;
    using Jurassic.Library;

    using JurassicExposer.FluentIL;

    public class TestClass : ITestClass
    {
        private readonly ExposerV2 exposer;
        public string GetValue(string input)
        {
            string current = (string)exposer.Unwrap(typeof(string), input);
            return null;
        }

        public object GetValue(object input)
        {
            var current = exposer.Unwrap(typeof(int), input);
            return null;
        }

        int ITestClass.Sum(int num1, int num2, int num3)
        {
            return num1 + num2 + num3;
        }

        public int Multiply(int num1, int num2)
        {
            return num1 * num2;
        }

        public MyJsonObject TestJson(MyJsonObject obj)
        {
            return new MyJsonObject(){ Name =  obj.Value, Value = obj.Value };
        }
    }

    public class TestClassProxy
    {
        private readonly ExposerV2 exposer;

        private readonly ScriptEngine engine;

        private readonly ITestClass realInstance;

        private readonly ILog log;

        public TestClassProxy(ExposerV2 exposer, ScriptEngine engine, ITestClass realInstance, ILog log)
        {
            this.exposer = exposer;
            this.engine = engine;
            this.realInstance = realInstance;
            this.log = log;
        }

        public string GetValue(string val)
        {
            string ret = (string)this.exposer.Unwrap(typeof(string), val);
            return this.realInstance.GetValue(ret);
        }

        [JSFunction(Name = "mul", NonStandard = false, Deprecated = false, Length = -1, IsWritable = true, IsEnumerable = false, IsConfigurable = true, Flags = JSFunctionFlags.None)]
        public int mul(int num, int num2)
        {
            int num3 = (int)this.exposer.Unwrap(typeof(int), num);
            int num4 = (int)this.exposer.Unwrap(typeof(int), num2);
            return (int)this.exposer.Wrap(typeof(int), null, this.realInstance.Multiply(num3, num4));
        }

        [JSFunction(Name = "json")]
        public ObjectInstance TestJson(ObjectInstance obj)
        {
            MyJsonObject o = (MyJsonObject)this.exposer.Unwrap(typeof(JSONObject), obj);

            return (ObjectInstance)this.exposer.Wrap(typeof(ObjectInstance), null, this.realInstance.TestJson(o));
        }

        public void Warn(string input)
        {
            string input1 = (string)this.exposer.Unwrap(typeof(string), input);
            this.log.Warn(input1);
        }
    }

    public interface IContainsCallback
    {
        string GetData(string input, Action<string, string> callback);
    }

    public interface ILog
    {
        [JSFunction(Name =  "debug")]
        void Debug(string line);

        [JSFunction(Name = "warn")]
        void Warn(string line);
    }

    public class ConsoleLog : ILog
    {
        public void Debug(string line)
        {
            Console.WriteLine(line);
        }

        public void Warn(string line)
        {
            Console.WriteLine(line);
        }
    }

    public class MyJsonObject
    {
        [JSProperty(Name = "name")]
        public string Name { get; set; }

        [JSProperty(Name = "value")]
        public string Value { get; set; }
    }

    public interface ITestClass
    {
        string GetValue(string input);

        [JSFunction(Name = "sum")]
        int Sum(int num1, int num2, int num3);

        [JSFunction(Name = "mul")]
        int Multiply(int num1, int num2);

        [JSFunction(Name = "json")]
        MyJsonObject TestJson(MyJsonObject obj);
    }

    class Program
    {
        static void Main(string[] args)
        {
            var eg = new ScriptEngine();
            var exposer = new ExposerV2();
            var dynAssInfo = IL.NewAssembly("testAssembly.dll");
            exposer.ExposeInterface<ITestClass>(dynAssInfo);
            eg.SetGlobal<ITestClass, TestClass>(exposer, "calc", new TestClass(), dynAssInfo);
            var logT = exposer.ExposeInterface<ILog>(dynAssInfo);
            dynAssInfo.Save();
            
            eg.SetGlobalValue("console", logT(eg, new ConsoleLog()));
            Stopwatch sw = Stopwatch.StartNew();
            for (var i = 0; i < 10; i++)
            {
                eg.Execute("var s = calc.sum(1, 2, 3); console.warn(s.toString());var jsobj = calc.json({name: '1', value: '2'}); console.warn('old name: ' + jsobj.name); jsobj.name = 'rafael nicoletti'; console.warn('name:' + jsobj.name);");
                Console.WriteLine("Elapsed: {0}", sw.ElapsedMilliseconds);
                sw.Restart();
            }
            Console.ReadLine();
        }
    }

    public static class ScriptEngineExtensions
    {
        public static void SetGlobal<TInterface, TObj>(this ScriptEngine engine, ExposerV2 v2, string name, TObj instance, DynamicAssemblyInfo dynAssInfo = null) where TObj : TInterface
        {
            var ctor = v2.ExposeInterface<TInterface>(dynAssInfo);
            engine.SetGlobalValue(name, ctor(engine, instance));
        }
    }
}
