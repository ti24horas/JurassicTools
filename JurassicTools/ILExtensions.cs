namespace JurassicTools
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;
    using System.Reflection.Emit;

    using Jurassic;
    using Jurassic.Library;

    public struct IlFieldResult
    {
        public FieldInfo RealInstance { get; set; }

        public FieldInfo ExposerInstance { get; set; }

        public Type InstanceType { get; set; }
    }

    public static class IlExtensions
    {
        private static void PopulateFieldFromConstructor(this ILGenerator gen, FieldInfo field, int argIndex)
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

        public static void CreateConstructor(this TypeBuilder builder, IlFieldResult fieldResult)
        {
            var ctorBuilder = builder.DefineConstructor(MethodAttributes.Public | MethodAttributes.HideBySig, CallingConventions.HasThis, new[] { typeof(ScriptEngine), typeof(JurassicExposer), fieldResult.InstanceType });

            #region ctor

            var ctorGen = ctorBuilder.GetILGenerator();
            ctorGen.Emit(OpCodes.Ldarg_0); // # this
            ctorGen.Emit(OpCodes.Ldarg_1); // > engine
            ctorGen.Emit(OpCodes.Call, ReflectionCache.ObjectInstance__ctor__ScriptEngine); // #:base(<)
            ctorGen.Emit(OpCodes.Ldarg_0); // # this
            #endregion

            ctorGen.PopulateFieldFromConstructor(fieldResult.ExposerInstance, 2);
            ctorGen.PopulateFieldFromConstructor(fieldResult.RealInstance, 3);

            #region PopulateFunctions
            ctorGen.Emit(OpCodes.Ldtoken, builder); // >[__this__]
            ctorGen.Emit(OpCodes.Call, ReflectionCache.Type__GetTypeFromHandle__RuntimeTypeHandle); // > typeof(<[__this__])
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
                proxyInst.SetReturnType(exposer.GetConvertOrWrapType(method.ReturnType));
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
                        exposer.EmitConvertOrUnwrap(methodGen, i, fieldResult.ExposerInstance, parameterInfos[i].ParameterType);
                    }
                }
                methodGen.Emit(method.IsVirtual ? OpCodes.Callvirt : OpCodes.Call, method); // >? <.Method(<*)

                exposer.EmitConvertOrWrap(methodGen, method.ReturnType);
                methodGen.Emit(OpCodes.Ret);
            }
        }
    }
}