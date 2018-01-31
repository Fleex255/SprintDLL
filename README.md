# SprintDLL

Write scripts that call native Win32 functions!

`rundll32` is often used to call arbitrary functions from the command line, but that [doesn't actually work](https://superuser.com/q/1074587/380318) - it assumes the called functions have its expected signature and doesn't parse parameter lists. SprintDLL supports user-supplied parameter lists and several calling conventions. You can use SprintDLL to call practically any native Windows function. Since SprintDLL can execute multiple instructions in one run, you can store data returned by a function and pass it to other functions.

## Documentation

Please see the [wiki](https://github.com/Fleex255/SprintDLL/wiki) for usage guidance.

## Download

You can download a compiled version from the [releases section](https://github.com/Fleex255/SprintDLL/releases).