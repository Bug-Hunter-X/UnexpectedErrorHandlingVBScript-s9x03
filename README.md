# Unexpected Error Handling in VBScript

This repository demonstrates a common, yet subtle, error in VBScript's error handling.  VBScript's implicit error handling can lead to unexpected runtime behavior if not carefully managed.

The `bug.vbs` file showcases a function that uses `Err.Raise` to explicitly throw a custom error. However, this approach requires careful attention to error handling within calling functions to prevent unexpected termination.

The `bugSolution.vbs` file illustrates a more robust approach to error handling, using `On Error Resume Next` and checking the `Err` object's properties to provide more informative and gracefully handled error conditions.