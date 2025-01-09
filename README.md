This repository demonstrates a subtle bug in VBScript related to the IsEmpty function and variant numeric types.  The bug occurs when IsEmpty is used to check if a variant variable holds a number. In certain situations, IsEmpty may return True even when the variable holds a numeric value, leading to unexpected behavior in the code. The solution shows how to correctly check for empty values while accounting for numeric variant types.