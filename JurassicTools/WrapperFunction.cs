using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using Jurassic;
using Jurassic.Library;

namespace JurassicTools
{
  internal class WrapperFunction : FunctionInstance
  {
      private readonly JurassicExposer exposer;

      private readonly Delegate _delegate;

    public WrapperFunction(JurassicExposer exposer, ScriptEngine engine, Delegate dele)
      : base(engine)
    {
        this.exposer = exposer;
        _delegate = dele;
    }

      public override object CallLateBound(object thisObject, params object[] argumentValues)
    {
      object[] args = new object[argumentValues.Length];
      ParameterInfo[] parameterInfos = _delegate.GetType().GetMethod("Invoke").GetParameters();
      Type[] types = parameterInfos.Select(pi => pi.ParameterType).ToArray();
      for (int i = 0; i < argumentValues.Length; i++)
      {
        args[i] = this.exposer.ConvertOrUnwrapObject(argumentValues[i], types[i]);
      }
      object ret = _delegate.DynamicInvoke(args);
      return this.exposer.ConvertOrWrapObject(ret);
    }
  }
}
