using System;
using System.Reflection;
using Ninject;
using Ninject.Extensions.NamedScope;
using Ninject.Parameters;

namespace GV.Ninject
{
    public static class Ioc
    {
        private static readonly MethodInfo _doGenericActionMethod = typeof(Ioc).GetMethod("DoGenericAction", BindingFlags.Static | BindingFlags.NonPublic);
        private static IKernel CurrentKernel { get; set; }

        public static void Initialize(IKernel kernel)
        {
            CurrentKernel = kernel;
        }

        public static T Resolve<T>()
        {
            return CurrentKernel.Get<T>();
        }

        public static T Resolve<T>(params IParameter[] parameters)
        {
            return CurrentKernel.Get<T>(parameters);
        }

        public static void Do(Type serviceType, Action<object> action)
        {
            _doGenericActionMethod.MakeGenericMethod(serviceType).Invoke(null, new object[] {action});
        }

        //don't remove this.
        private static void DoGenericAction<T>(Action<T> action)
        {
            Do(action);
        }

        public static void Do<T>(Action<T> action)
        {
            var namedScopeParameter = new NamedScopeParameter("FPR.Ioc InCallScope");
            using (namedScopeParameter.Scope)
            {
                using (var proxy = Resolve<DisposeNotifyingProxy<T>>(namedScopeParameter))
                {
                    action(proxy.Service);
                }
            }
        }

        public class DisposeNotifyingProxy<T> : DisposeNotifyingObject
        {
            public DisposeNotifyingProxy(T service)
            {
                Service = service;
            }

            public T Service { get; }
        }
    }
}