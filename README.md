
# TOKEN Function

Excel add-in function that returns a single substring from a **string expression**, given custom **delimiters** and the substring **position**.

**Syntax**

**TOKEN( expression [, delimiters] [, position]  )**

The **TOKEN** function syntax has these arguments:

| **Argument** | **Description** |
| :------------------- | :------------------- |
| **expression**  | Required 1st parameter. String expression containing substrings and delimiters. Can be passed by value or by reference. If **expression** is an empty string (""), **TOKEN** returns an empty string. |
| **delimiters** | Optional 2nd or 3rd parameter. String composed of delimiter characters used to identify substring limits. **TOKEN** identifies within **expression** the longest sequences of characters from **delimiters** and uses these sequences to identify substrings. The string of **delimiters** can be passed by value (in which case surrounding quotation marks are ignored) or by reference (in which case surrounding quotation marks are included in the set of delimiter characters). If **delimiters** is omitted, the default is a string made of the space character (" "). If **delimiter** is an empty string (""), **TOKEN** returns the entire **expression** string.  |
| **position** | Optional 2nd or 3rd parameter. Number specifying the position of the substring to be returned by **TOKEN**, counting from 1. If **position** is omitted, the default is 1. Can be passed by value or by reference. If the 2nd parameter is numeric while the 3rd parameter is omitted or empty, or if the 3rd parameter is numeric while the 2nd parameter is empty, then that numeric parameter is interpreted as **position**. |
