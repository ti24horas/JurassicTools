namespace JurassicExposer.FluentIL
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    using System.Reflection;
    using System.Reflection.Emit;

    using global::FluentIL;
    using global::FluentIL.Infos;

    using Jurassic;
    using Jurassic.Library;

    using PropertyAttributes = Jurassic.Library.PropertyAttributes;

    public class ExposerV2
    {
        private readonly IDictionary<Type, Delegate> creatonalFunctions = new Dictionary<Type, Delegate>();

        public Func<ScriptEngine, T, object> ExposeInterface<T>(DynamicAssemblyInfo dynAssInfo = null)
        {
            if (!typeof(T).IsInterface)
            {
                throw new InvalidOperationException("type T should be an interface");
            }

            if (this.creatonalFunctions.ContainsKey(typeof(T)))
            {
                return (Func<ScriptEngine, T, object>)this.creatonalFunctions[typeof(T)];
            }
            var typeName = "Proxy_" + typeof(T).Name;
            DynamicTypeInfo t;
            
            t = dynAssInfo != null ? dynAssInfo.WithType(typeName) : IL.NewType(typeName);

            t = t.Inherits<ObjectInstance>();

            t.WithField("engine", typeof(ScriptEngine))
                .WithField("realInstance", typeof(T))
                .WithField("exposer", typeof(ExposerV2));

            var objectInstanceConstructor = typeof(ObjectInstance).GetConstructor(
                BindingFlags.NonPublic | BindingFlags.Instance,
                null,
                new[] { typeof(ScriptEngine) },
                null);

            t.WithConstructor(
                cd =>
                {
                    cd
                          .BodyDefinition()
                          .Ldarg(0, 2).Call(objectInstanceConstructor) // pass script engine
                          .Ldarg(0, 1).Stfld("exposer")
                          .Ldarg(0, 2).Stfld("engine")
                          .Ldarg(0, 3).Stfld("realInstance")
                          .Ldarg(0)
                          .Call(typeof(ObjectInstance).GetMethod("PopulateFunctions", BindingFlags.NonPublic | BindingFlags.Instance, null, new Type[0], null))
                          .Ret();
                }, typeof(ExposerV2), typeof(ScriptEngine), typeof(T));

            foreach (var mtd in typeof(T).GetMethods(BindingFlags.Public | BindingFlags.Instance))
            {
                this.ExposeMethod(t, mtd);
            }

            var type = t.Complete();

            return (this.creatonalFunctions[typeof(T)] = new Func<ScriptEngine, T, object>((eng, real) => Activator.CreateInstance(type, this, eng, real))) as Func<ScriptEngine, T, object>;
        }

        public static Type GetCompatibleType(Type t)
        {
            if (t == typeof(void))
            {
                return t;
            }

            switch (Type.GetTypeCode(t))
            {
                case TypeCode.Boolean: return typeof(BooleanInstance);
                case TypeCode.Byte: return typeof(int);
                case TypeCode.Char:
                    return typeof(string);
                case TypeCode.DateTime:
                    return typeof(DateInstance);
                case TypeCode.Decimal:
                    return typeof(double);
                case TypeCode.Int16:
                case TypeCode.Int32:
                    return typeof(int);
                case TypeCode.Int64:
                    return typeof(long);
                case TypeCode.String:
                    return typeof(string);
                default:
                    return typeof(ObjectInstance);
            }
        }

        public object Unwrap(Type parameterType, object input)
        {
            if (!(input is ObjectInstance))
            {
                return input;
            }

            var inst1 = (ObjectInstance)input;
            var properties = from x in parameterType.GetProperties().Where(a => a.CanWrite)
                             select new { prop = x, val = this.Unwrap(x.PropertyType, inst1.GetPropertyValue(x.Name)) };
            var inst2 = Activator.CreateInstance(parameterType);
            foreach (var p in properties)
            {
                p.prop.SetValue(inst2, p.val, null);
            }
            return inst2;
        }

        private class ObjectGetterFunction : FunctionInstance
        {
            private readonly ExposerV2 v2;

            private readonly object realInstance;

            private readonly PropertyInfo p;

            public ObjectGetterFunction(ExposerV2 v2, ScriptEngine engine, object realInstance, PropertyInfo p)
                : base(engine)
            {
                this.v2 = v2;
                this.realInstance = realInstance;
                this.p = p;
            }

            public override object CallLateBound(object thisObject, params object[] argumentValues)
            {
                return this.v2.Wrap(ExposerV2.GetCompatibleType(this.p.PropertyType), this.Engine, this.p.GetValue(this.realInstance, null));
            }
        }

        private class ObjectSetterFunction : FunctionInstance
        {
            private readonly ExposerV2 v2;

            private readonly object input;

            private readonly PropertyInfo p;

            public ObjectSetterFunction(ExposerV2 v2, ScriptEngine engine, object input, PropertyInfo p)
                : base(engine)
            {
                this.v2 = v2;
                this.input = input;
                this.p = p;
            }

            public override object CallLateBound(object thisObject, params object[] argumentValues)
            {
                this.p.SetValue(this.input, this.v2.Unwrap(this.p.PropertyType, argumentValues[0]), null);
                return argumentValues[0];
            }
        }

        public object Wrap(Type type, ScriptEngine engine, object input)
        {
            if (type == typeof(ObjectInstance))
            {
                var properties = from x in input.GetType().GetProperties().Where(a => a.CanWrite)
                                 select new { prop = x, 
                                     name = x.GetCustomAttributes(typeof(JSPropertyAttribute), true).OfType<JSPropertyAttribute>().Select(a=>a.Name).FirstOrDefault() ?? x.Name,
                                     val = this.Wrap(x.PropertyType, engine, x.GetValue(input)) };

                var obj = engine.Object.Construct();
                foreach (var p in properties)
                {
                    var descriptor = new PropertyDescriptor(new ObjectGetterFunction(this, engine, input, p.prop), new ObjectSetterFunction(this, engine, input, p.prop), PropertyAttributes.FullAccess);
                    obj.DefineProperty(p.name, descriptor, false);
                }
                return obj;
                // TODO: check if is a list of objects

                // TODO: create a objectinstance from input
            }
            return input;
        }

        private void ExposeMethod(DynamicTypeInfo dynamicTypeInfo, MethodInfo mtd)
        {
            var infoAttributes = mtd.GetCustomAttributes(typeof(JSFunctionAttribute), true).OfType<JSFunctionAttribute>().FirstOrDefault()
                                 ?? new JSFunctionAttribute() { Name = mtd.Name };

            var method = dynamicTypeInfo.WithMethod(infoAttributes.Name);
            method.TurnOffAttributes(MethodAttributes.Virtual);

            method.Returns(GetCompatibleType(mtd.ReturnType));

            var properties = (from x in typeof(JSFunctionAttribute).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                              let val = x.GetValue(infoAttributes, null)
                              where x.CanWrite
                              where val != null
                              select new { prop = x, vale = val }).ToArray();

            foreach (var p in mtd.GetParameters())
            {
                // TODO: check if parameter type is javascript compatible
                method.WithParameter(GetCompatibleType(p.ParameterType), p.Name);
                method.WithVariable(p.ParameterType);
            }

            foreach (var p in mtd.GetParameters())
            {
                method.Body.Ldarg(0).Ldfld("exposer")
                    .Emit(OpCodes.Ldtoken, p.ParameterType)
                    .Call(typeof(Type).GetMethod("GetTypeFromHandle", BindingFlags.Static | BindingFlags.Public));

                method.Body.Ldarg((uint)(p.Position + 1));
                if (p.ParameterType.IsValueType)
                {
                    method.Body.Box(p.ParameterType);
                }
                method.Body.Callvirt(typeof(ExposerV2).GetMethod("Unwrap", BindingFlags.Public | BindingFlags.Instance, null, new Type[] { typeof(Type), typeof(object) }, null));
                if (p.ParameterType.IsValueType)
                {
                    method.Body.UnboxAny(p.ParameterType);
                }
                else
                {
                    method.Body.Emit(OpCodes.Castclass, p.ParameterType);
                }

                method.Body.Stloc((uint)p.Position);
            }

            if (method.ReturnType != typeof(void))
            {
                method.Body.Ldarg(0)
                   .Ldfld("exposer")
                   .Emit(OpCodes.Ldtoken, method.ReturnType)
                   .Call(typeof(Type).GetMethod("GetTypeFromHandle", BindingFlags.Static | BindingFlags.Public));
                method.Body.Ldarg(0).Ldfld("engine");
            }


            // TODO: check parameter equality
            method.Body.Ldarg(0)
                .Ldfld("realInstance")
                .Ldloc(Enumerable.Range(0, mtd.GetParameters().Length).Select(a => (uint)a).ToArray())
                .Callvirt(mtd);


            if (method.ReturnType != typeof(void))
            {
                if (method.ReturnType.IsValueType)
                {
                    method.Body.Box(method.ReturnType);
                }
                method.Body.Callvirt(
                    typeof(ExposerV2).GetMethod(
                        "Wrap",
                        BindingFlags.Public | BindingFlags.Instance,
                        null,
                        new[] { typeof(Type), typeof(ScriptEngine), typeof(object) },
                        null));
                if (method.ReturnType.IsValueType)
                {
                    method.Body.UnboxAny(method.ReturnType);
                }
                else
                {
                    method.Body.Emit(OpCodes.Castclass, GetCompatibleType(method.ReturnType));
                }
            }
            method.Body.Ret();

            var jsAttributeBuilder = new CustomAttributeBuilder(typeof(JSFunctionAttribute).GetConstructor(new Type[0]), new object[0], properties.Select(a => a.prop).ToArray(), properties.Select(a => a.vale).ToArray());
            method.MethodBuilder.SetCustomAttribute(jsAttributeBuilder);
        }
    }
}
