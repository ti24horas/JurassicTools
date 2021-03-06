﻿namespace JurassicTools
{
    using System;
    using System.Collections;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.ComponentModel;
    using System.Linq;
    using System.Reflection;
    using System.Reflection.Emit;
    using System.Threading;

    using Jurassic;
    using Jurassic.Library;

    public class JurassicExposer
    {
        private readonly ScriptEngine engine;

        private readonly AssemblyBuilder MyAssembly;
        private readonly ModuleBuilder MyModule;

        private readonly Dictionary<Type, JurassicInfo[]> TypeInfos = new Dictionary<Type, JurassicInfo[]>();

        private readonly HashSet<Type> staticProxyCache = new HashSet<Type>();
        private readonly Dictionary<Type, Type> InstanceProxyCache = new Dictionary<Type, Type>();

        private long DelegateCounter;

        private readonly Dictionary<Tuple<Type, WeakReference>, Delegate> DelegateProxyCache = new Dictionary<Tuple<Type, WeakReference>, Delegate>();
        private readonly Dictionary<long, object> DelegateFunctions = new Dictionary<long, object>();

        public JurassicExposer(ScriptEngine engine)
        {
            this.engine = engine;
            var name = new AssemblyName("JurassicProxy");
            MyAssembly = AppDomain.CurrentDomain.DefineDynamicAssembly(name, AssemblyBuilderAccess.RunAndSave);
            MyModule = MyAssembly.DefineDynamicModule(string.Format("{0}.dll", name.Name), string.Format("{0}.dll", name.Name));
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

        internal JurassicInfo[] FindInfos(Type type)
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

        public void ExposeClass<T>(String name = null)
        {
            ExposeClass(typeof(T), this.engine, name);
        }

        public void ExposeClass(Type typeT, ScriptEngine engine, String name = null)
        {
            if (this.staticProxyCache.Contains(typeT))
            {
                return;
            }

            if (name == null) name = typeT.Name;
            JurassicInfo[] infos = FindInfos(typeT);

            Type proxiedType;
            // public class JurassicStaticProxy.T : ClrFunction
            TypeBuilder typeBuilder = MyModule.DefineType("JurassicStaticProxy." + typeT.FullName, TypeAttributes.Class | TypeAttributes.Public, typeof(ClrFunction));

            // public .ctor(ScriptEngine engine, plainString name)
            // : base(engine.Function.InstancePrototype, name, engine.Object)
            // base.PopulateFunctions(null, BindingFlags.Public | BindingFlags.Static /*| BindingFlags.DeclaredOnly*/);
            // base.PopulateFields(null);
            var exposer = typeBuilder.DefineField("exposer", typeof(JurassicExposer), FieldAttributes.Private | FieldAttributes.InitOnly);
            var ctorBuilder = typeBuilder.DefineConstructor(MethodAttributes.Public | MethodAttributes.HideBySig, CallingConventions.HasThis, new[] { typeof(ScriptEngine), typeof(JurassicExposer), typeof(string) });
            ctorBuilder.CreateConstructor(null,
                gen =>
                {
                    gen.Emit(OpCodes.Ldarg_0);
                    gen.Emit(OpCodes.Ldarg_1);
                    gen.Emit(OpCodes.Callvirt, ReflectionCache.ScriptEngine__get_Function); // > <.Function
                    gen.Emit(OpCodes.Callvirt, ReflectionCache.FunctionInstance__get_InstancePrototype); // > <.InstancePrototype
                    gen.Emit(OpCodes.Ldarg_3);
                    gen.Emit(OpCodes.Ldarg_1);
                    gen.Emit(OpCodes.Callvirt, ReflectionCache.ScriptEngine__get_Object);
                    gen.Emit(OpCodes.Call, ReflectionCache.ClrFunction__ctor__ObjectInstance_String_ObjectInstance);
                    gen.PopulateFieldFromConstructor(exposer, 2);

                });

            if (typeT.IsEnum)
            {
                if (Attribute.IsDefined(typeT, typeof(FlagsAttribute)))
                {
                    Type enumType = Enum.GetUnderlyingType(typeT);
                    foreach (string v in Enum.GetNames(typeT))
                    {
                        FieldBuilder field = typeBuilder.DefineField(v, enumType, FieldAttributes.Public | FieldAttributes.Static | FieldAttributes.Literal);
                        field.SetConstant(Convert.ChangeType(Enum.Parse(typeT, v), enumType));
                        field.SetCustomAttribute(new CustomAttributeBuilder(ReflectionCache.JSField__ctor, new object[] { }));
                    }
                }
                else
                {
                    foreach (string v in Enum.GetNames(typeT))
                    {
                        FieldBuilder field = typeBuilder.DefineField(v, typeof(string), FieldAttributes.Public | FieldAttributes.Static | FieldAttributes.Literal);
                        field.SetConstant(v);
                        field.SetCustomAttribute(new CustomAttributeBuilder(ReflectionCache.JSField__ctor, new object[] { }));
                    }
                }
            }
            else
            {
                // [JsConstructorFunction]
                // public ObjectInstance Construct(params object[] args)
                // return JurassicExposer.WrapObject(Activator.CreateInstance(typeof(T), args), Engine);
                var jsctorBuilder = typeBuilder.DefineMethod("Construct", MethodAttributes.Public | MethodAttributes.HideBySig, CallingConventions.HasThis, typeof(ObjectInstance), new[] { typeof(Object[]) });

                jsctorBuilder.SetCustomAttribute(new CustomAttributeBuilder(ReflectionCache.JSConstructorFunctionAttribute__ctor, new object[] { }));
                var jsctorParams = jsctorBuilder.DefineParameter(1, ParameterAttributes.None, "args");
                jsctorParams.SetCustomAttribute(new CustomAttributeBuilder(ReflectionCache.ParamArrayAttribute__ctor, new object[] { }));
                var jsctorGen = jsctorBuilder.GetILGenerator();
                jsctorGen.Emit(OpCodes.Ldarg_0);
                jsctorGen.Emit(OpCodes.Ldfld, exposer);

                jsctorGen.Emit(OpCodes.Ldarg_0);
                jsctorGen.Emit(OpCodes.Ldfld, exposer);

                jsctorGen.Emit(OpCodes.Ldtoken, typeT); // T

                jsctorGen.Emit(OpCodes.Call, ReflectionCache.Type__GetTypeFromHandle__RuntimeTypeHandle); // > typeof(T)

                jsctorGen.Emit(OpCodes.Ldarg_1); // > args

                jsctorGen.Emit(OpCodes.Callvirt, typeof(JurassicExposer).GetMethod("CreateNewInstance", new[] { typeof(Type), typeof(object[]) }));

                jsctorGen.Emit(OpCodes.Call, ReflectionCache.JurassicExposer__WrapObject__Object_ScriptEngine); // > JurassicExposer.WrapObject(<, <)
                jsctorGen.Emit(OpCodes.Ret);
            }

            proxiedType = typeBuilder.CreateType();
            var proxiedInstance = (ClrFunction)Activator.CreateInstance(proxiedType, engine, this, name);
            engine.SetGlobalValue(name, proxiedInstance);
            this.staticProxyCache.Add(typeT);
        }

        public void ExposeInstance(object instance, String name)
        {
            object inst = ConvertOrWrapObject(instance);
            engine.SetGlobalValue(name, inst);
        }

        public object CreateInstanceObject(object instance)
        {
            return ConvertOrWrapObject(instance);
        }


        public object ConvertOrWrapObject(object instance)
        {
            if (instance == null) return null;
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
            if (instance == null) return null;
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
                var arr = (Array)instance;
                object[] arr2 = new object[arr.Length];
                for (int i = 0; i < arr.Length; i++)
                {
                    var value = arr.GetValue(i);
                    if (value == null)
                    {
                        arr2[i] = null;
                    }
                    else
                    {
                        arr2[i] = ConvertOrUnwrapObject(value, value.GetType().GetConvertOrUnwrapType());
                    }
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
                                                           nameAndValue => ConvertOrUnwrapObject(nameAndValue.Value, nameAndValue.Value.GetType().GetConvertOrUnwrapType()));
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
                        var arr = instance as ArrayInstance;
                        var actualObj = type.GetGenericArguments().First();
                        var list = Array.CreateInstance(actualObj, arr.Length);
                        for (var i = 0; i < arr.Length; i++)
                        {
                            if (arr[i] is ObjectInstance && Type.GetTypeCode(arr[i].GetType()) == TypeCode.Object)
                            {
                                var inst = this.ConvertOrUnwrapObject(arr[i], actualObj);
                                list.SetValue(inst, i);
                            }
                        }
                        return list;
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
                                              let jsProp = x.GetCustomAttributes(typeof(JSPropertyAttribute), true).OfType<JSPropertyAttribute>().FirstOrDefault()
                                              select
                                                  new
                                                      {
                                                          propertyName = x.Name,
                                                          jsPropertyName = (jsProp != null ? jsProp.Name : x.Name) ?? x.Name,
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
                                    if (property.property.PropertyType.IsEnum)
                                    {
                                        if (value is string)
                                        {
                                            property.property.SetValue(targetObject, Enum.Parse(property.property.PropertyType, value.ToString()), null);
                                        }
                                    }
                                    else
                                    {
                                        var convertedValue = Convert.ChangeType(value, property.property.PropertyType);
                                        property.property.SetValue(targetObject, convertedValue, null);
                                    }

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
                il.EmitConvertOrWrap(parameterInfos[i].ParameterType, localFunction);
                if (parameterInfos[i].ParameterType.IsValueType)
                {
                    il.Emit(OpCodes.Box, parameterInfos[i].ParameterType);
                }

                il.Emit(OpCodes.Stelem_Ref); // <par[<i] = <
            }

            il.Emit(OpCodes.Ldnull); // >[thisObject]
            il.Emit(OpCodes.Ldloc, par); // >par
            il.Emit(OpCodes.Callvirt, ReflectionCache.FunctionInstance__CallLateBound__Object_aObject); // >? <localFunction.CallLateBound(<[thisObject], <par)
            if (mi.ReturnType == typeof(void))
            {
                il.Emit(OpCodes.Pop);
            }
            else
            {
                il.EmitConvertOrWrap(mi.ReturnType);
            }
            il.Emit(OpCodes.Ret);
            Delegate dele = dm.CreateDelegate(delegateType);
            return dele;
        }

        public object ConvertObject(object instance, Type type)
        {
            return ConvertOrUnwrapObject(instance, type);
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
                var fieldCreator = typeBuilder.PopulateFields(type);
                var ctorBuilder = typeBuilder.DefineConstructor(MethodAttributes.Public | MethodAttributes.HideBySig, CallingConventions.HasThis, new[] { typeof(ScriptEngine), typeof(JurassicExposer), type });

                ctorBuilder.CreateConstructor(fieldCreator);
                typeBuilder.CreateMethods(this, fieldCreator);

                // public ... Property
                PropertyInfo[] piInsts = type.GetProperties(BindingFlags.Public | BindingFlags.Instance /*| BindingFlags.DeclaredOnly*/);
                foreach (PropertyInfo piInst in piInsts)
                {
                    Attribute[] infoAttributes = GetAttributes(infos, piInst.Name, typeof(JSPropertyAttribute));
                    if (!Attribute.IsDefined(piInst, typeof(JSPropertyAttribute)) && infoAttributes.Length == 0) continue;
                    MethodInfo piInstGet = piInst.GetGetMethod();
                    MethodInfo piInstSet = piInst.GetSetMethod();
                    if (piInstGet == null && piInstSet == null) continue;
                    PropertyBuilder proxyInstance = typeBuilder.DefineProperty(piInst.Name, piInst.Attributes, piInst.PropertyType.GetConvertOrWrapType(), null);
                    proxyInstance.CopyCustomAttributesFrom(piInst, infoAttributes);
                    if (piInstGet != null /*&& !methodNames.Contains(piInstGet.Name)*/)
                    {
                        //methodNames.Add(piInstGet.Name);
                        MethodBuilder proxyInstanceGet = typeBuilder.DefineMethod(piInstGet.Name, piInstGet.Attributes);
                        proxyInstanceGet.SetReturnType(piInstGet.ReturnType.GetConvertOrWrapType());
                        proxyInstanceGet.CopyParametersFrom(this, piInstGet);
                        proxyInstanceGet.CopyCustomAttributesFrom(piInstGet);
                        ILGenerator getGen = proxyInstanceGet.GetILGenerator();

                        getGen.Emit(OpCodes.Ldarg_0); // # this
                        getGen.Emit(OpCodes.Ldfld, fieldCreator.RealInstance); // > #.realInstance
                        ParameterInfo[] parameterInfos = piInstGet.GetParameters();
                        for (int i = 0; i < parameterInfos.Length; i++)
                        {
                            getGen.EmitConvertOrUnwrap(i, fieldCreator.ExposerInstance, parameterInfos[i].ParameterType);
                        }
                        getGen.Emit(piInstGet.IsVirtual ? OpCodes.Callvirt : OpCodes.Call, piInstGet); // <.Property <*
                        getGen.EmitConvertOrWrap(piInstGet.ReturnType);
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
                        setGen.Emit(OpCodes.Ldfld, fieldCreator.RealInstance); // > #.realInstance
                        var parameterInfos = piInstSet.GetParameters();
                        for (int i = 0; i < parameterInfos.Length; i++)
                        {
                            setGen.EmitConvertOrUnwrap(i, fieldCreator.ExposerInstance, parameterInfos[i].ParameterType);
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
                    ilAdd.Emit(OpCodes.Ldfld, fieldCreator.RealInstance); // > #.realInstance
                    ParameterInfo[] parameterInfos = eiAdd.GetParameters();
                    for (int i = 0; i < parameterInfos.Length; i++)
                    {
                        ilAdd.EmitConvertOrUnwrap(i, fieldCreator.ExposerInstance, parameterInfos[i].ParameterType);
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
                    ilRemove.Emit(OpCodes.Ldfld, fieldCreator.RealInstance); // > #.realInstance
                    parameterInfos = eiRemove.GetParameters();
                    for (int i = 0; i < parameterInfos.Length; i++)
                    {
                        ilRemove.EmitConvertOrUnwrap(i, fieldCreator.ExposerInstance, parameterInfos[i].ParameterType);
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

        public Attribute[] GetAttributes(JurassicInfo[] infos, string name, Type attributeType = null)
        {
            return (infos == null || infos.Length == 0)
                     ? new Attribute[0]
                     : infos.Where(i => String.Equals(i.MemberName, name))
                            .SelectMany(i => i.Attributes.Where(a => attributeType == null || a.GetType() == attributeType))
                            .ToArray();
        }

        public object CreateNewInstance(Type objectType, object[] args)
        {
            return Activator.CreateInstance(objectType, args);
        }
    }
}
