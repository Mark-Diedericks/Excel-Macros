from System.Dynamic import DynamicObject as _do
from System import Func, Array, Object
from System.Runtime.CompilerServices import CallSite
from Microsoft.CSharp.RuntimeBinder import (
        Binder, CSharpBinderFlags, CSharpArgumentInfoFlags,
        RuntimeBinderException, CSharpArgumentInfo
        )


class _Injected:
    def __dir__(self):
        return list(self.GetDynamicMemberNames())

    def _create_dyn_getter(self, name):
        arg_info = Array.CreateInstance(CSharpArgumentInfo, 1)
        arg_info[0] = CSharpArgumentInfo.Create(
                getattr(CSharpArgumentInfoFlags, 'None'), None
                )

        binder = Binder.GetMember(
                getattr(CSharpBinderFlags, 'None'),
                name,
                self.GetType(),
                arg_info
                )

        callsite = CallSite[Func[CallSite, Object, Object]].Create(binder)
        return lambda obj: callsite.Target(callsite, obj)

    def __getattr__(self, name):
        cache_name = '__Pyxos_callsite_cache'
        cls = self.__class__

        if not hasattr(cls, cache_name):
            setattr(cls, cache_name, {})

        cache = getattr(cls, cache_name)
        if name not in cache:
            cache[name] = self._create_dyn_getter(name)

        try:
            return cache[name](self)
        except RuntimeBinderException:
            raise AttributeError('Cannot access dynamic attribute ' + name)


def _patch():
    _do.__bases__ = (_Injected, Object)