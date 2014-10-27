using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Threading;
using Jurassic;
using Jurassic.Library;

namespace JurassicTools
{
    using System.ComponentModel;
    using System.Net.NetworkInformation;
    using System.Security.Cryptography;

    public class JurassicExposer
    {
        private readonly ScriptEngine engine;

        private readonly AssemblyBuilder MyAssembly;
        private readonly ModuleBuilder MyModule;

        private readonly Dictionary<Type, JurassicInfo[]> TypeInfos = new Dictionary<Type, JurassicInfo[]>();

        private readonly Dictionary<Type, Type> StaticProxyCache = new Dictionary<Type, Type>();
        private readonly Dictionary<Type, Type> InstanceProxyCache = new Dictionary<Type, Type>();

        private long DelegateCounter;

        private readonly Dictionary<Tuple<Type, WeakReference>, Delegate> DelegateProxyCache = new Dictionary<Tuple<Type, WeakReference>, Delegate>();
        private readonly Dictionary<long, object> DelegateFunctions = new Dictionary<long, object>();

        public JurassicExposer(ScriptEngine engine, string assemblyName)
        {
            this.engine = engine;
            AssemblyName name = new AssemblyName(assemblyName);
            MyAssembly = AppDomain.CurrentDomain.DefineDynamicAssembly(name, AssemblyBuilderAccess.RunAndSave);
            MyModule = MyAssembly.DefineDynamicModule(string.Format("{0}.dll", assemblyName), string.Format("{0}.dll", assemblyName));
        }

        public void SaveAssembly()
        {
            MyAssembly.Save(MyAssembly.GetName().Name + ".dll");
        }

        internal object GetFunction(long index)
        {
            return DelegateFunctions[index];
        }

        private long AddFunction(Type type, object function)
        {
            long l = Interlocked.Increment(ref DelegateCounter);
            DelegateFunctions[l] = function;
            return l;
        }

        public void RegisterInfos<T>(params JurassicInfo[] infos)
        {
            RegisterInfos(typeof(T), infos);
        }

        public void RegisterInfos(Type typeT, params JurassicInfo[] infos)
        {
            if (TypeInfos.ContainsKey(typeT)) return;
            TypeInfos[typeT] = infos;
        }

        private JurassicInfo[] FindInfos(Type type)
        {
            List<JurassicInfo> infos = new List<JurassicInfo>();
            Type t = type;
            while (t != null && t != typeof(Object))
            {
                if (TypeInfos.ContainsKey(t))
                {
                    infos.AddRange(TypeInfos[t].Where(ni => !infos.Any(i => String.Equals(i.MemberName, ni.MemberName))));
                }
                t = t.BaseType;
            }
            foreach (Type implementedInterface in type.GetInterfaces())
            {
                if (TypeInfos.ContainsKey(implementedInterface))
                {
                    infos.AddRange(TypeInfos[implementedInterface].Where(ni => !infos.Any(i => String.Equals(i.MemberName, ni.MemberName))));
                }
            }
            return infos.ToArray();
        }
        /*
        public void ExposeClass<T>(String name = null)
        {
            ExposeClass(typeof(T), name);
        }
        */

        public void ExposeInstance(object instance, String name)
        {
            object inst = ConvertOrWrapObject(instance);
            engine.SetGlobalValue(name, inst);
        }

        public object CreateInstanceObject(object instance)
        {
            return ConvertOrWrapObject(instance);
        }

        public Type GetConvertOrWrapType(Type type)
        {
            if (type == typeof(void)) return typeof(void);
            if (type == typeof(ScriptEngine)) return typeof(ScriptEngine); // JSFunction with HasEngineParameter
            if (type.IsEnum)
            {
                return Attribute.IsDefined(type, typeof(FlagsAttribute)) ? Enum.GetUnderlyingType(type) : typeof(string);
            }
            if (type.IsArray)
            {
                return typeof(ArrayInstance);
            }
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Boolean:
                    return typeof(bool);
                case TypeCode.Byte:
                    return typeof(int);
                case TypeCode.Char:
                    return typeof(string);
                case TypeCode.DateTime:
                    return typeof(DateInstance);
                case TypeCode.Decimal:
                    return typeof(double);
                case TypeCode.Double:
                    return typeof(double);
                case TypeCode.Int16:
                    return typeof(int);
                case TypeCode.Int32:
                    return typeof(int);
                case TypeCode.Int64:
                    return typeof(double);
                case TypeCode.Object:
                    return typeof(ObjectInstance);
                case TypeCode.SByte:
                    return typeof(int);
                case TypeCode.Single:
                    return typeof(double);
                case TypeCode.String:
                    return typeof(string);
                case TypeCode.UInt16:
                    return typeof(int);
                case TypeCode.UInt32:
                    return typeof(uint);
                case TypeCode.UInt64:
                    return typeof(double);
                default:
                    throw new ArgumentException(string.Format("Cannot convert value of type {0}.", type), "type");
            }
        }

        public Type GetConvertOrUnwrapType(Type type)
        {
            if (type == typeof(void)) return typeof(void);
            if (type == typeof(ConcatenatedString)) return typeof(string);
            if (type == typeof(ScriptEngine)) return typeof(ScriptEngine); // JSFunction with HasEngineParameter
            if (type == typeof(ArrayInstance)) return typeof(object[]);
            if (type == typeof(BooleanInstance) || type == typeof(bool)) return typeof(bool);
            if (type == typeof(NumberInstance) || type == typeof(Byte) || type == typeof(Decimal) || type == typeof(Double) || type == typeof(Int16) ||
                type == typeof(Int32) || type == typeof(Int64) || type == typeof(SByte) || type == typeof(Single) || type == typeof(UInt16) || type == typeof(UInt32) ||
                type == typeof(UInt64)) return typeof(double);
            if (type == typeof(StringInstance) || type == typeof(string)) return typeof(string);
            if (type == typeof(DateInstance)) return typeof(DateTime);
            if (type == typeof(FunctionInstance)) return typeof(Delegate); // TODO?
            if (type == typeof(ObjectInstance)) return typeof(IDictionary<string, object>);
            throw new ArgumentException("unkown type for unwrap: " + type.FullName);
        }

        public object ConvertOrWrapObject(object instance)
        {
            if (instance == null) return Undefined.Value;
            //if (instance == null) return engine.Object.InstancePrototype;
            Type type = instance.GetType();
            if (type == typeof(void)) return null;
            if (type == typeof(ScriptEngine)) return instance;
            if (type.IsEnum)
            {
                return Attribute.IsDefined(type, typeof(FlagsAttribute)) ? Convert.ChangeType(instance, Enum.GetUnderlyingType(type)) : Enum.GetName(type, instance);
            }
            if (type.IsArray)
            {
                Array arr = (Array)instance;
                ArrayInstance arr2 = engine.Array.New();
                for (int i = 0; i < arr.Length; i++)
                {
                    arr2[i] = ConvertOrWrapObject(arr.GetValue(i));
                }
                return arr2;
            }
            if (type == typeof(JSONString))
            {
                return JSONObject.Parse(engine, ((JSONString)instance).plainString);
            }
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Boolean:
                    return (bool)instance;
                case TypeCode.Byte:
                    return (int)(byte)instance;
                case TypeCode.Char:
                    return new string((char)instance, 1);
                case TypeCode.DateTime:
                    DateTime dt = (DateTime)instance;
                    if (dt == DateTime.MinValue) return null;
                    return engine.Date.Construct((((DateTime)instance).ToUniversalTime().Ticks - new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero).Ticks) / 10000.0);
                case TypeCode.Decimal:
                    return decimal.ToDouble((decimal)instance);
                case TypeCode.Double:
                    return (double)instance;
                case TypeCode.Int16:
                    return (int)(short)instance;
                case TypeCode.Int32:
                    return (int)instance;
                case TypeCode.Int64:
                    return (double)(long)instance;
                case TypeCode.Object:
                    if (instance is FunctionInstance)
                    {
                        /*var oldDelegates = DelegateProxyCache.Where(kvp => !kvp.Key.IsAlive);
                        foreach (KeyValuePair<WeakReference, Delegate> keyValuePair in oldDelegates)
                        {
                          DelegateProxyCache.Remove(keyValuePair.Key);
                        }
                        Delegate dele;
                        if (DelegateProxyCache.Any(pair => pair.Key.Target == instance))
                        {
                          dele = DelegateProxyCache.First(pair => pair.Key.Target == instance).Value;
                        }
                        else
                        {
                          //dele = objects => ((FunctionInstance)instance).Call(null, objects);
                          dele = null;
                          DelegateProxyCache[new WeakReference(instance)] = dele;
                        }
                        return dele;*/
                    }
                    if (instance is ObjectInstance) return instance;
                    IDictionary dictionary = instance as IDictionary;
                    if (dictionary != null)
                    {
                        ObjectInstance obj = WrapObject(instance);
                        foreach (DictionaryEntry dictionaryEntry in dictionary)
                        {
                            obj.SetPropertyValue(dictionaryEntry.Key.ToString(), ConvertOrWrapObject(dictionaryEntry.Value), true);
                        }
                        return obj;
                    }
                    if (typeof(NameValueCollection).IsAssignableFrom(type))
                    {
                        ObjectInstance obj = WrapObject(instance);
                        NameValueCollection nvc = (NameValueCollection)instance;
                        foreach (var item in nvc.AllKeys)
                        {
                            string[] vals = nvc.GetValues(item);
                            if (vals.Length == 1) obj.SetPropertyValue(item, vals[0], true);
                            else
                            {
                                ArrayInstance arr = engine.Array.Construct(vals);
                                obj.SetPropertyValue(item, arr, true);
                            }
                        }
                        return obj;
                    }
                    if (type.GetInterfaces().Any(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>)))
                    {
                        ArrayInstance arr = engine.Array.Construct();
                        IEnumerable ienum = (IEnumerable<Object>)instance;
                        int i = 0;
                        foreach (object item in ienum)
                        {
                            arr.Push(ConvertOrWrapObject(item));
                        }
                        return arr;
                    }
                    if (type.GetInterfaces().Any(i => i == typeof(IEnumerable)))
                    {
                        ArrayInstance arr = engine.Array.Construct();
                        IEnumerable ienum = (IEnumerable)instance;
                        int i = 0;
                        foreach (object item in ienum)
                        {
                            arr.Push(ConvertOrWrapObject(item));
                        }
                        return arr;
                    }
                    return WrapObject(instance);
                case TypeCode.SByte:
                    return (int)(sbyte)instance;
                case TypeCode.Single:
                    return (double)(float)instance;
                case TypeCode.String:
                    return instance;
                case TypeCode.UInt16:
                    return (int)(ushort)instance;
                case TypeCode.UInt32:
                    return (uint)instance;
                case TypeCode.UInt64:
                    return (double)(ulong)instance;
                default:
                    throw new ArgumentException(string.Format("Cannot store value of type {0}.", type), "instance");
            }
        }

        public object ConvertOrUnwrapObject(object instance, Type type)
        {
            if (instance == null) return Null.Value;
            //Type type = instance.GetType();
            if (type == typeof(void)) return null;
            if (instance is ConcatenatedString)
            {
                return (instance as ConcatenatedString).ToString();
            }
            if (type.IsEnum)
            {
                if (!Attribute.IsDefined(type, typeof(FlagsAttribute)))
                {
                    return Enum.Parse(type, (string)instance);
                }
            }
            if (type.IsArray)
            {
                ArrayInstance arr = (ArrayInstance)instance;
                object[] arr2 = new object[arr.Length];
                for (int i = 0; i < arr.Length; i++)
                {
                    arr2[i] = ConvertOrUnwrapObject(arr[i], GetConvertOrUnwrapType(arr[i] == null ? null : arr[i].GetType()));
                }
                return arr2;
            }
            if (type == typeof(JSONString))
            {
                throw new NotImplementedException("TODO: Unwrap JSONString");
            }
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Boolean:
                    return (bool)Convert.ChangeType(instance, type);
                case TypeCode.Byte:
                    return (byte)Convert.ChangeType(instance, type);
                case TypeCode.Char:
                    string str = instance as string;
                    return string.IsNullOrEmpty(str) ? char.MinValue : str[0];
                case TypeCode.DateTime:
                    return new DateTime((long)((((DateInstance)instance).GetTime() * 10000) + new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero).Ticks)).ToLocalTime(/* TODO: check*/);
                case TypeCode.Decimal:
                    return new decimal((double)Convert.ChangeType(instance, type));
                case TypeCode.Double:
                    return (double)Convert.ChangeType(instance, type);
                case TypeCode.Int16:
                    return (short)Convert.ChangeType(instance, type);
                case TypeCode.Int32:
                    return (int)Convert.ChangeType(instance, type);
                case TypeCode.Int64:
                    return (long)Convert.ChangeType(instance, type);
                case TypeCode.Object:
                    if (instance is FunctionInstance && typeof(Delegate).IsAssignableFrom(type))
                    {
                        var oldDelegates = DelegateProxyCache.Where(kvp => !kvp.Key.Item2.IsAlive);
                        foreach (KeyValuePair<Tuple<Type, WeakReference>, Delegate> oldDelegate in oldDelegates)
                        {
                            DelegateProxyCache.Remove(oldDelegate.Key);
                        }

                        FunctionInstance function = (FunctionInstance)instance;
                        if (DelegateProxyCache.Any(kvp => kvp.Key.Item1 == type && kvp.Key.Item2.Target == function))
                        {
                            return DelegateProxyCache.First(kvp => kvp.Key.Item1 == type && kvp.Key.Item2.Target == function).Value;
                        }

                        WeakReference weakReference = new WeakReference(function);
                        Tuple<Type, WeakReference> tuple = new Tuple<Type, WeakReference>(type, weakReference);
                        var dele = UnwrapFunction(type, function);
                        DelegateProxyCache.Add(tuple, dele);
                        DelegateProxyCache[tuple] = dele;
                        return dele;
                    }
                    if (typeof(IDictionary<string, object>).IsAssignableFrom(type))
                    {
                        ObjectInstance obj = instance as ObjectInstance;
                        if (obj == null) return null;
                        return obj.Properties.ToDictionary(nameAndValue => nameAndValue.Name,
                                                           nameAndValue => ConvertOrUnwrapObject(nameAndValue.Value, GetConvertOrUnwrapType(nameAndValue.Value.GetType())));
                    }
                    if (typeof(NameValueCollection).IsAssignableFrom(type))
                    {
                        /* NameValueCollection nvc = (NameValueCollection)instance;
                         foreach (var item in nvc.AllKeys)
                         {
                           var val = nvc[item];
                           obj.SetPropertyValue(item, ConvertOrWrapObject(val, engine), true);
                         }*/
                        ObjectInstance obj = instance as ObjectInstance;
                        if (obj == null) return null;
                        NameValueCollection nvc = new NameValueCollection();
                        foreach (PropertyNameAndValue propertyNameAndValue in obj.Properties)
                        {
                            string key = propertyNameAndValue.Name;
                            object val = propertyNameAndValue.Value;
                            if (val is ArrayInstance)
                            {
                                ArrayInstance arr = (ArrayInstance)val;
                                for (int i = 0; i < arr.Length; i++)
                                {
                                    if (arr[i] is ObjectInstance) continue;
                                    if (arr[i] is Null) nvc.Add(key, null);
                                    else nvc.Add(key, arr[i].ToString());
                                }
                            }
                            else if (val is ObjectInstance) continue;
                            else if (val is Null) nvc.Add(key, null);
                            else nvc.Add(key, val.ToString());
                        }
                        return nvc;
                    }
                    //if (instance is ObjectInstance) return instance;
                    if (type.GetInterfaces().Any(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>)))
                    {
                        /*IEnumerable ienum = (IEnumerable<Object>)instance;
                        int i = 0;
                        foreach (var item in ienum)
                        {
                          obj[i++] = ConvertOrWrapObject(item, engine);
                        }
                        obj["length"] = i;*/
                        throw new NotImplementedException("TODO: Unwrap IEnumerable<>");
                    }
                    if (type.GetInterfaces().Any(i => i == typeof(IEnumerable)))
                    {
                        /*IEnumerable ienum = (IEnumerable)instance;
                        int i = 0;
                        foreach (var item in ienum)
                        {
                          obj[i++] = ConvertOrWrapObject(item, engine);
                        }
                        obj["length"] = i;*/
                        throw new NotImplementedException("TODO: Unwrap IEnumerable");
                    }
                    Type instanceType = instance.GetType();
                    if (instance is FunctionInstance)
                    {
                        return instanceType.FullName.StartsWith("JurassicInstanceProxy.")
                                   ? instanceType.GetField("realInstance").GetValue(instance)
                                   : instance;
                    }
                    if (instance is ObjectInstance)
                    {
                        var targetObject = Activator.CreateInstance(type);
                        var allJsProperties = from x in type.GetProperties()
                                              let jsProperty =
                                                  x.GetCustomAttributes(true)
                                                  .OfType<JSPropertyAttribute>()
                                                  .FirstOrDefault()
                                              where jsProperty != null
                                              select
                                                  new
                                                      {
                                                          propertyName = x.Name,
                                                          jsPropertyName = jsProperty.Name ?? x.Name,
                                                          property = x
                                                      };

                        allJsProperties = allJsProperties.ToList();
                        var objInstance = (instance as ObjectInstance);
                        foreach (var property in allJsProperties)
                        {
                            var value = objInstance[property.jsPropertyName] ?? objInstance[property.propertyName];
                            if (value == null)
                            {
                                var @default =
                                    property.property.GetCustomAttributes(typeof(DefaultValueAttribute), true)
                                        .OfType<DefaultValueAttribute>()
                                        .FirstOrDefault();
                                if (@default != null)
                                {
                                    value = @default.Value;
                                }
                            }
                            if (value != null)
                            {

                                if (property.property.PropertyType == typeof(string))
                                {
                                    value = value.ToString();
                                }
                                try
                                {
                                    property.property.SetValue(targetObject, value, null);
                                }
                                catch (ArgumentException)
                                {
                                    var convertedValue = Convert.ChangeType(value, property.property.PropertyType);
                                    property.property.SetValue(targetObject, convertedValue, null);

                                }

                            }
                        }
                        return targetObject;
                    }
                    return instanceType.FullName.StartsWith("JurassicInstanceProxy.") ? instanceType.GetField("realInstance").GetValue(instance) : instance;
                case TypeCode.SByte:
                    return (sbyte)instance;
                case TypeCode.Single:
                    return (float)instance;
                case TypeCode.String:
                    return instance;
                case TypeCode.UInt16:
                    return (ushort)instance;
                case TypeCode.UInt32:
                    return (uint)instance;
                case TypeCode.UInt64:
                    return (ulong)instance;
                default:
                    throw new ArgumentException(string.Format("Cannot store value of type {0}.", type), "instance");
            }
        }

        private Delegate UnwrapFunction(Type delegateType, FunctionInstance function)
        {
            if (!typeof(Delegate).IsAssignableFrom(delegateType)) return null;
            MethodInfo mi = delegateType.GetMethod("Invoke");
            ParameterInfo[] parameterInfos = mi.GetParameters();
            DynamicMethod dm = new DynamicMethod("DynamicMethod_for_" + delegateType.Name, mi.ReturnType, parameterInfos.Select(pi => pi.ParameterType).ToArray(),
                                                 typeof(JurassicExposer));
            ILGenerator il = dm.GetILGenerator();
            var localFunction = il.DeclareLocal(typeof(FunctionInstance));
            var par = il.DeclareLocal(typeof(object[]));
            long l = AddFunction(delegateType, function);
            il.Emit(OpCodes.Ldc_I8, l); // >[index]
            il.Emit(OpCodes.Call, ReflectionCache.JurassicExposer__GetFunction__long); // >[function] JurassicExposer.GetFunction(<[index])
            il.Emit(OpCodes.Stloc, localFunction); // localFunction = <[function]
            il.Emit(OpCodes.Ldloc, localFunction); // >localFunction
            il.Emit(OpCodes.Ldc_I4, parameterInfos.Length); // >[length]
            il.Emit(OpCodes.Newarr, typeof(object)); // > new object[<[length]]
            il.Emit(OpCodes.Stloc, par); // par = <
            for (int i = 0; i < parameterInfos.Length; i++)
            {
                il.Emit(OpCodes.Ldloc, par); // >par
                il.Emit(OpCodes.Ldc_I4, i); // >i
                il.Emit(OpCodes.Ldarg, i); // > arg*
                if (!Attribute.IsDefined(parameterInfos[i], typeof(ParamArrayAttribute))) EmitConvertOrWrap(il, parameterInfos[i].ParameterType, null, localFunction);
                if (parameterInfos[i].ParameterType.IsValueType) il.Emit(OpCodes.Box, parameterInfos[i].ParameterType);
                il.Emit(OpCodes.Stelem_Ref); // <par[<i] = <
            }
            il.Emit(OpCodes.Ldnull); // >[thisObject]
            il.Emit(OpCodes.Ldloc, par); // >par
            il.Emit(OpCodes.Callvirt, ReflectionCache.FunctionInstance__CallLateBound__Object_aObject); // >? <localFunction.CallLateBound(<[thisObject], <par)
            if (mi.ReturnType == typeof(void)) il.Emit(OpCodes.Pop);
            else EmitConvertOrWrap(il, mi.ReturnType, null);
            il.Emit(OpCodes.Ret);
            Delegate dele = dm.CreateDelegate(delegateType);
            return dele;
        }

        private void EmitConvertOrWrap(ILGenerator gen, Type type, FieldBuilder fldExposer, LocalBuilder localFunction = null)
        {
            if (type == typeof(void)) return;
            // > value (from caller)
            if (type.IsValueType)
            {
                gen.Emit(OpCodes.Box, type);
            }
            if (localFunction == null)
            {
            }
            else gen.Emit(OpCodes.Ldloc, localFunction); // >localFunction
            
            gen.Emit(OpCodes.Callvirt, ReflectionCache.JurassicExposer__ConvertOrWrapObject__Object_ScriptEngine); // JurassicExposer.ConvertOrWrapObject(<, <, <)
            Type convertOrWrapType = GetConvertOrWrapType(type);
            if (convertOrWrapType.IsValueType) gen.Emit(OpCodes.Unbox_Any, convertOrWrapType);
        }

        private void EmitConvertOrUnwrap(ILGenerator gen, int parameterIndex, FieldInfo exposer, Type type)
        {
            if (type == typeof(void)) return;

            // > value (from caller)
            Type convertOrWrapType = GetConvertOrWrapType(type);
            //if (convertOrWrapType.IsValueType) gen.Emit(OpCodes.Box, convertOrWrapType);
            gen.Emit(OpCodes.Ldarg_0);
            gen.Emit(OpCodes.Ldfld, exposer);
            gen.Emit(OpCodes.Ldarg, parameterIndex + 1);
            
            if (type.IsValueType)
            {
                if (type.IsEnum)
                {
                    //return Attribute.IsDefined(type, typeof(FlagsAttribute)) ? Convert.ChangeType(instance, Enum.GetUnderlyingType(type)) : Enum.GetName(type, instance);
                    if (Attribute.IsDefined(type, typeof(FlagsAttribute))) gen.Emit(OpCodes.Box, Enum.GetUnderlyingType(type));
                    else gen.Emit(OpCodes.Box, typeof(string));
                }
                else
                {
                    // TODO: needed?
                    gen.Emit(OpCodes.Box, type); // why? bugs!
                }
            }
            
            Type realType = type;
            if (type.IsByRef || type.IsPointer) realType = type.GetElementType();
            gen.Emit(OpCodes.Ldtoken, realType); // > type

            gen.Emit(OpCodes.Call, ReflectionCache.Type__GetTypeFromHandle__RuntimeTypeHandle); // > typeof(<)

            gen.Emit(OpCodes.Callvirt, ReflectionCache.JurassicExposer__ConvertOrUnwrapObject__Object_Static); // JurassicExposer.ConvertOrUnwrapObject(<, <)
            if (type.IsValueType) gen.Emit(OpCodes.Unbox_Any, type);
        }

        public object ConvertObject(object instance, Type type)
        {
            return ConvertOrUnwrapObject(instance, type);
            return instance;
        }

        public ObjectInstance WrapObject(object instance)
        {
            Type type = instance.GetType();
            JurassicInfo[] infos = FindInfos(type);

            Type proxiedType;
            if (!InstanceProxyCache.TryGetValue(type, out proxiedType))
            {
                // public class JurassicInstanceProxy.T : ObjectInstance
                TypeBuilder typeBuilder = MyModule.DefineType("JurassicInstanceProxy." + type.FullName, TypeAttributes.Class | TypeAttributes.Public,
                                                               typeof(ObjectInstance));

                // public object realInstance
                FieldBuilder fldInstance = typeBuilder.DefineField("realInstance", instance.GetType(), FieldAttributes.Private);
                var fieldExposer = typeBuilder.DefineField("exposer", typeof(JurassicExposer), FieldAttributes.Private);

                // .ctor(ScriptEngine engine, object instance)
                // : base(engine)
                // base.PopulateFunctions(typeof(__this__), BindingFlags.Public | BindingFlags.Instance /*| BindingFlags.DeclaredOnly*/);
                // realInstance = instance;
                ConstructorBuilder ctorBuilder = typeBuilder.DefineConstructor(MethodAttributes.Public | MethodAttributes.HideBySig, CallingConventions.HasThis,
                                                                               new[] { typeof(ScriptEngine), typeof(JurassicExposer), instance.GetType() });
                ILGenerator ctorGen = ctorBuilder.GetILGenerator();

                ctorGen.Emit(OpCodes.Ldarg_0); // # this
                ctorGen.Emit(OpCodes.Ldarg_1); // > engine
                ctorGen.Emit(OpCodes.Call, ReflectionCache.ObjectInstance__ctor__ScriptEngine); // #:base(<)
                ctorGen.Emit(OpCodes.Ldarg_0); // # this
                ctorGen.Emit(OpCodes.Ldtoken, typeBuilder); // >[__this__]
                ctorGen.Emit(OpCodes.Call, ReflectionCache.Type__GetTypeFromHandle__RuntimeTypeHandle); // > typeof(<[__this__])
                ctorGen.Emit(OpCodes.Ldc_I4, (int)(BindingFlags.Public | BindingFlags.Instance /*| BindingFlags.DeclaredOnly*/)); // > flags
                ctorGen.Emit(OpCodes.Call, ReflectionCache.ObjectInstance__PopulateFunctions__Type_BindingFlags); // #.PopulateFunctions(<, <);
                ctorGen.Emit(OpCodes.Ldarg_0);
                ctorGen.Emit(OpCodes.Ldarg_2);
                ctorGen.Emit(OpCodes.Stfld, fieldExposer); // #.realInstance = <

                ctorGen.Emit(OpCodes.Ldarg_0); // # this
                ctorGen.Emit(OpCodes.Ldarg_3); // > instance
                ctorGen.Emit(OpCodes.Stfld, fldInstance); // #.realInstance = <
                ctorGen.Emit(OpCodes.Ret);

                // public ... Method(...)
                MethodInfo[] miInsts = type.GetMethods(BindingFlags.Public | BindingFlags.Instance /*| BindingFlags.DeclaredOnly*/);
                List<string> methodNames = new List<string>();
                foreach (MethodInfo miInst in miInsts)
                {
                    
                    if (methodNames.Contains(miInst.Name)) continue;
                    else methodNames.Add(miInst.Name);
                    Attribute[] infoAttributes = GetAttributes(infos, miInst.Name, typeof(JSFunctionAttribute));
                    if (!Attribute.IsDefined(miInst, typeof(JSFunctionAttribute)) && infoAttributes.Length == 0) continue;
                    MethodBuilder proxyInst = typeBuilder.DefineMethod(miInst.Name, miInst.Attributes);
                    proxyInst.SetReturnType(GetConvertOrWrapType(miInst.ReturnType));
                    proxyInst.CopyParametersFrom(this, miInst);
                    proxyInst.CopyCustomAttributesFrom(miInst, infoAttributes);
                    ILGenerator methodGen = proxyInst.GetILGenerator();

                    ParameterInfo[] parameterInfos = miInst.GetParameters();
                    if (miInst.ReturnType != typeof(void))
                    {
                        methodGen.Emit(OpCodes.Ldarg_0);
                        methodGen.Emit(OpCodes.Ldfld, fieldExposer);
                    }
                    methodGen.Emit(OpCodes.Ldarg_0);
                    methodGen.Emit(OpCodes.Ldfld, fldInstance);
                    for (int i = 0; i < parameterInfos.Length; i++)
                    {
                        //methodGen.Emit(OpCodes.Ldarg, i + 1); // > arg*
                        if (!Attribute.IsDefined(parameterInfos[i], typeof(ParamArrayAttribute))) this.EmitConvertOrUnwrap(methodGen, i, fieldExposer, parameterInfos[i].ParameterType);
                    }
                    methodGen.Emit(miInst.IsVirtual ? OpCodes.Callvirt : OpCodes.Call, miInst); // >? <.Method(<*)

                    EmitConvertOrWrap(methodGen, miInst.ReturnType, fieldExposer);
                    methodGen.Emit(OpCodes.Ret);
                }

                // public ... Property
                PropertyInfo[] piInsts = type.GetProperties(BindingFlags.Public | BindingFlags.Instance /*| BindingFlags.DeclaredOnly*/);
                foreach (PropertyInfo piInst in piInsts)
                {
                    Attribute[] infoAttributes = GetAttributes(infos, piInst.Name, typeof(JSPropertyAttribute));
                    if (!Attribute.IsDefined(piInst, typeof(JSPropertyAttribute)) && infoAttributes.Length == 0) continue;
                    MethodInfo piInstGet = piInst.GetGetMethod();
                    MethodInfo piInstSet = piInst.GetSetMethod();
                    if (piInstGet == null && piInstSet == null) continue;
                    PropertyBuilder proxyInstance = typeBuilder.DefineProperty(piInst.Name, piInst.Attributes, GetConvertOrWrapType(piInst.PropertyType), null);
                    proxyInstance.CopyCustomAttributesFrom(piInst, infoAttributes);
                    if (piInstGet != null /*&& !methodNames.Contains(piInstGet.Name)*/)
                    {
                        //methodNames.Add(piInstGet.Name);
                        MethodBuilder proxyInstanceGet = typeBuilder.DefineMethod(piInstGet.Name, piInstGet.Attributes);
                        proxyInstanceGet.SetReturnType(GetConvertOrWrapType(piInstGet.ReturnType));
                        proxyInstanceGet.CopyParametersFrom(this, piInstGet);
                        proxyInstanceGet.CopyCustomAttributesFrom(piInstGet);
                        ILGenerator getGen = proxyInstanceGet.GetILGenerator();

                        getGen.Emit(OpCodes.Ldarg_0); // # this
                        getGen.Emit(OpCodes.Ldfld, fldInstance); // > #.realInstance
                        ParameterInfo[] parameterInfos = piInstGet.GetParameters();
                        for (int i = 0; i < parameterInfos.Length; i++)
                        {
                            //getGen.Emit(OpCodes.Ldarg, i + 1); // > arg*
                            if (!Attribute.IsDefined(parameterInfos[i], typeof(ParamArrayAttribute))) this.EmitConvertOrUnwrap(getGen, i, fieldExposer, parameterInfos[i].ParameterType);
                        }
                        getGen.Emit(piInstGet.IsVirtual ? OpCodes.Callvirt : OpCodes.Call, piInstGet); // <.Property <*
                        EmitConvertOrWrap(getGen, piInstGet.ReturnType, fieldExposer);
                        getGen.Emit(OpCodes.Ret);
                        proxyInstance.SetGetMethod(proxyInstanceGet);
                    }
                    if (piInstSet != null /*&& !methodNames.Contains(piInstSet.Name)*/)
                    {
                        //methodNames.Add(piInstSet.Name);
                        MethodBuilder proxyInstanceSet = typeBuilder.DefineMethod(piInstSet.Name, piInstSet.Attributes);
                        proxyInstanceSet.SetReturnType(piInstSet.ReturnType);
                        proxyInstanceSet.CopyParametersFrom(this, piInstSet);
                        proxyInstanceSet.CopyCustomAttributesFrom(piInstSet);
                        ILGenerator setGen = proxyInstanceSet.GetILGenerator();

                        setGen.Emit(OpCodes.Ldarg_0); // # this
                        setGen.Emit(OpCodes.Ldfld, fldInstance); // > #.realInstance
                        ParameterInfo[] parameterInfos = piInstSet.GetParameters();
                        for (int i = 0; i < parameterInfos.Length; i++)
                        {
                            //setGen.Emit(OpCodes.Ldarg, i + 1); // > arg*
                            if (!Attribute.IsDefined(parameterInfos[i], typeof(ParamArrayAttribute))) this.EmitConvertOrUnwrap(setGen, i, fieldExposer, parameterInfos[i].ParameterType);
                        }
                        setGen.Emit(piInstSet.IsVirtual ? OpCodes.Callvirt : OpCodes.Call, piInstSet); // <.Property = <*
                        setGen.Emit(OpCodes.Ret);
                        proxyInstance.SetSetMethod(proxyInstanceSet);
                    }
                }

                // public event ...
                EventInfo[] eiInsts = type.GetEvents(BindingFlags.Public | BindingFlags.Instance /*| BindingFlags.DeclaredOnly*/);
                foreach (EventInfo eventInfo in eiInsts)
                {
                    Attribute[] infoAttributes = GetAttributes(infos, eventInfo.Name, typeof(JSEventAttribute));
                    if (!Attribute.IsDefined(eventInfo, typeof(JSEventAttribute)) && infoAttributes.Length == 0) continue;
                    JSEventAttribute eventAttribute = (JSEventAttribute)Attribute.GetCustomAttribute(eventInfo, typeof(JSEventAttribute));
                    if (eventAttribute == null) eventAttribute = (JSEventAttribute)infoAttributes.FirstOrDefault(a => a is JSEventAttribute);
                    MethodInfo eiAdd = eventInfo.GetAddMethod();
                    if (eventAttribute == null) eventAttribute = new JSEventAttribute();
                    string addName = eventAttribute.AddPrefix + (eventAttribute.Name ?? eventInfo.Name);
                    MethodBuilder proxyAdd = typeBuilder.DefineMethod(addName, eiAdd.Attributes);
                    proxyAdd.SetReturnType(eiAdd.ReturnType);
                    proxyAdd.CopyParametersFrom(this, eiAdd);
                    proxyAdd.CopyCustomAttributesFrom(eiAdd, new JSFunctionAttribute());
                    ILGenerator ilAdd = proxyAdd.GetILGenerator();
                    ilAdd.Emit(OpCodes.Ldarg_0); // # this
                    ilAdd.Emit(OpCodes.Ldfld, fldInstance); // > #.realInstance
                    ParameterInfo[] parameterInfos = eiAdd.GetParameters();
                    for (int i = 0; i < parameterInfos.Length; i++)
                    {
                        //ilAdd.Emit(OpCodes.Ldarg, i + 1); // > arg*
                        if (!Attribute.IsDefined(parameterInfos[i], typeof(ParamArrayAttribute))) this.EmitConvertOrUnwrap(ilAdd, i, fieldExposer, parameterInfos[i].ParameterType);
                    }
                    ilAdd.Emit(eiAdd.IsVirtual ? OpCodes.Callvirt : OpCodes.Call, eiAdd);
                    ilAdd.Emit(OpCodes.Ret);

                    MethodInfo eiRemove = eventInfo.GetRemoveMethod();
                    string removeName = eventAttribute.RemovePrefix + (eventAttribute.Name ?? eventInfo.Name);
                    MethodBuilder proxyRemove = typeBuilder.DefineMethod(removeName, eiRemove.Attributes);
                    proxyRemove.SetReturnType(eiRemove.ReturnType);
                    proxyRemove.CopyParametersFrom(this, eiRemove);
                    proxyRemove.CopyCustomAttributesFrom(eiRemove, new JSFunctionAttribute());
                    ILGenerator ilRemove = proxyRemove.GetILGenerator();
                    ilRemove.Emit(OpCodes.Ldarg_0); // # this
                    ilRemove.Emit(OpCodes.Ldfld, fldInstance); // > #.realInstance
                    parameterInfos = eiRemove.GetParameters();
                    for (int i = 0; i < parameterInfos.Length; i++)
                    {
                        //ilRemove.Emit(OpCodes.Ldarg, i + 1); // > arg*
                        if (!Attribute.IsDefined(parameterInfos[i], typeof(ParamArrayAttribute))) this.EmitConvertOrUnwrap(ilRemove, i, fieldExposer, parameterInfos[i].ParameterType);
                    }
                    ilRemove.Emit(eiRemove.IsVirtual ? OpCodes.Callvirt : OpCodes.Call, eiRemove);
                    ilRemove.Emit(OpCodes.Ret);
                }

                proxiedType = typeBuilder.CreateType();
                InstanceProxyCache[type] = proxiedType;
            }

            ObjectInstance proxiedInstance = (ObjectInstance)Activator.CreateInstance(proxiedType, this.engine, this, instance);
            return proxiedInstance;
        }

        private static Attribute[] GetAttributes(JurassicInfo[] infos, string name, Type attributeType = null)
        {
            return (infos == null || infos.Length == 0)
                     ? new Attribute[0]
                     : infos.Where(i => String.Equals(i.MemberName, name))
                            .SelectMany(i => i.Attributes.Where(a => attributeType == null || a.GetType() == attributeType))
                            .ToArray();
        }
    }
}
