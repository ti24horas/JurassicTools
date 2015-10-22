namespace JurassicTools
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Reflection.Emit;

    using Jurassic;
    using Jurassic.Library;

    public class IlFieldResult
    {
        public FieldInfo RealInstance { get; set; }

        public FieldInfo ExposerInstance { get; set; }

        public Type InstanceType { get; set; }
    }

    public static class IlExtensions
    {
        public static void PopulateFieldFromConstructor(this ILGenerator gen, FieldInfo field, int argIndex)
        {
            gen.Emit(OpCodes.Ldarg_0);
            switch (argIndex)
            {
                case 1:
                    gen.Emit(OpCodes.Ldarg_1);
                    break;
                case 2:
                    gen.Emit(OpCodes.Ldarg_2);
                    break;

                case 3:
                    gen.Emit(OpCodes.Ldarg_3);
                    break;
                case 0:
                    throw new InvalidOperationException("argument 0 is always 'this' instance");
                default:
                    gen.Emit(OpCodes.Ldarg, argIndex);
                    break;
            }

            gen.Emit(OpCodes.Stfld, field); // #.Exposer = <
        }

        public static IlFieldResult PopulateFields(this TypeBuilder typeBuilder, Type instanceType)
        {
            var fldInstance = typeBuilder.DefineField("realInstance", instanceType, FieldAttributes.Private | FieldAttributes.InitOnly);
            var fieldExposer = typeBuilder.DefineField("exposer", typeof(JurassicExposer), FieldAttributes.Private | FieldAttributes.InitOnly);
            return new IlFieldResult { ExposerInstance = fieldExposer, RealInstance = fldInstance, InstanceType = instanceType };
        }

        public static void CreateConstructor(this ConstructorBuilder builder, IlFieldResult fieldResult, Action<ILGenerator> onEmitArgs = null, int baseCtorArgumentCount = 1)
        {
            #region ctor

            var ctorGen = builder.GetILGenerator();
            if (onEmitArgs != null)
            {
                onEmitArgs(ctorGen);
            }
            else
            {
                ctorGen.Emit(OpCodes.Ldarg_0); // # this
                ctorGen.Emit(OpCodes.Ldarg_1); // > engine
                ctorGen.Emit(OpCodes.Call, ReflectionCache.ObjectInstance__ctor__ScriptEngine); // #:base(<)
            }

            ctorGen.Emit(OpCodes.Ldarg_0); // # this
            #endregion

            if (fieldResult != null)
            {
                ctorGen.PopulateFieldFromConstructor(fieldResult.ExposerInstance, 2);
                ctorGen.PopulateFieldFromConstructor(fieldResult.RealInstance, 3);
            }

            #region PopulateFunctions
            ctorGen.Emit(OpCodes.Ldarg_0); // >[__this__]
            ctorGen.Emit(OpCodes.Callvirt, typeof(Type).GetMethod("GetType", new Type[0])); // >[__this__]

            //ctorGen.Emit(OpCodes.Call, ReflectionCache.Type__GetTypeFromHandle__RuntimeTypeHandle); // > typeof(<[__this__])
            ctorGen.Emit(OpCodes.Ldc_I4, (int)(BindingFlags.Public | BindingFlags.Instance /*| BindingFlags.DeclaredOnly*/)); // > flags
            ctorGen.Emit(OpCodes.Callvirt, ReflectionCache.ObjectInstance__PopulateFunctions__Type_BindingFlags); // #.PopulateFunctions(<, <);
            #endregion

            #region populate fields
            ctorGen.Emit(OpCodes.Ldarg_0); // # this
            ctorGen.Emit(OpCodes.Ldnull); // > null
            ctorGen.Emit(OpCodes.Callvirt, ReflectionCache.ObjectInstance__PopulateFields__Type); // #.PopulateFields(<)
            #endregion

            ctorGen.Emit(OpCodes.Ret);
        }

        public static void CreateMethods(this TypeBuilder builder, JurassicExposer exposer, IlFieldResult fieldResult)
        {
            var methods = fieldResult.InstanceType.GetMethods(BindingFlags.Public | BindingFlags.Instance /*| BindingFlags.DeclaredOnly*/);
            var methodNames = new List<string>();
            var infos = exposer.FindInfos(fieldResult.InstanceType);
            foreach (var method in methods.Where(miInst => !methodNames.Contains(miInst.Name)))
            {
                methodNames.Add(method.Name);
                var infoAttributes = exposer.GetAttributes(infos, method.Name, typeof(JSFunctionAttribute));
                if (!Attribute.IsDefined(method, typeof(JSFunctionAttribute)) && infoAttributes.Length == 0)
                {
                    continue;
                }

                var proxyInst = builder.DefineMethod(method.Name, method.Attributes);
                proxyInst.SetReturnType(method.ReturnType.GetConvertOrWrapType());
                proxyInst.CopyParametersFrom(exposer, method);
                proxyInst.CopyCustomAttributesFrom(method, infoAttributes);
                var methodGen = proxyInst.GetILGenerator();

                var parameterInfos = method.GetParameters();
                if (method.ReturnType != typeof(void))
                {
                    methodGen.Emit(OpCodes.Ldarg_0);
                    methodGen.Emit(OpCodes.Ldfld, fieldResult.ExposerInstance);
                }

                methodGen.Emit(OpCodes.Ldarg_0);
                methodGen.Emit(OpCodes.Ldfld, fieldResult.RealInstance);
                for (var i = 0; i < parameterInfos.Length; i++)
                {
                    //methodGen.Emit(OpCodes.Ldarg, i + 1); // > arg*
                    if (!Attribute.IsDefined(parameterInfos[i], typeof(ParamArrayAttribute)))
                    {
                        methodGen.EmitConvertOrUnwrap(i, fieldResult.ExposerInstance, parameterInfos[i].ParameterType);
                    }
                }
                methodGen.Emit(method.IsVirtual ? OpCodes.Callvirt : OpCodes.Call, method); // >? <.Method(<*)

                methodGen.EmitConvertOrWrap(method.ReturnType);
                methodGen.Emit(OpCodes.Ret);
            }
        }

        internal static void EmitConvertOrUnwrap(this ILGenerator gen, int parameterIndex, FieldInfo exposer, Type type)
        {
            if (type == typeof(void))
            {
                return;
            }

            gen.Emit(OpCodes.Ldarg_0);
            gen.Emit(OpCodes.Ldfld, exposer);
            gen.Emit(OpCodes.Ldarg, parameterIndex + 1);

            if (type.IsValueType)
            {
                if (type.IsEnum)
                {
                    gen.Emit(OpCodes.Box, Attribute.IsDefined(type, typeof(FlagsAttribute)) ? Enum.GetUnderlyingType(type) : typeof(string));
                }
                else
                {
                    // TODO: needed?
                    gen.Emit(OpCodes.Box, type); // why? bugs!
                }
            }

            var realType = type;
            if (type.IsByRef || type.IsPointer)
            {
                realType = type.GetElementType();
            }

            gen.Emit(OpCodes.Ldtoken, realType); // > type

            gen.Emit(OpCodes.Call, ReflectionCache.Type__GetTypeFromHandle__RuntimeTypeHandle); // > typeof(<)

            gen.Emit(OpCodes.Callvirt, ReflectionCache.JurassicExposer__ConvertOrUnwrapObject__Object_Static); // JurassicExposer.ConvertOrUnwrapObject(<, <)
            if (type.IsValueType)
            {
                gen.Emit(OpCodes.Unbox_Any, type);
            }
        }

        internal static void EmitConvertOrWrap(this ILGenerator gen, Type type, LocalBuilder localFunction = null)
        {
            if (type == typeof(void))
            {
                return;
            }

            // > value (from caller)
            if (type.IsValueType)
            {
                gen.Emit(OpCodes.Box, type);
            }

            if (localFunction != null)
            {
                gen.Emit(OpCodes.Ldloc, localFunction); // >localFunction
            }

            gen.Emit(OpCodes.Callvirt, ReflectionCache.JurassicExposer__ConvertOrWrapObject__Object_ScriptEngine); // JurassicExposer.ConvertOrWrapObject(<, <, <)
            var convertOrWrapType = type.GetConvertOrWrapType();
            if (convertOrWrapType.IsValueType)
            {
                gen.Emit(OpCodes.Unbox_Any, convertOrWrapType);
            }
        }

        internal static Type GetConvertOrWrapType(this Type type)
        {
            if (type == typeof(void))
            {
                return typeof(void);
            }

            if (type == typeof(ScriptEngine))
            {
                return typeof(ScriptEngine); // JSFunction with HasEngineParameter
            }

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

        public static Type GetConvertOrUnwrapType(this Type type)
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

    }
}