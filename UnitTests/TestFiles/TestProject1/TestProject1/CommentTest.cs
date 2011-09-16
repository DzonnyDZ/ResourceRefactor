using System;
using System.Collections.Generic;
using System.Text;

namespace TestProject1
{
    /// <summary>
    /// File is used to find all instances of "Test String". There are 2 valid instances
    /// </summary>
    class CommentTest
    {
        string t = "Test String" + /* "Test String" */"Test String";
        string t2 = "Test Str";//"Test String"
        ////"Test String"//
        string t3 = "Test";
    }
}
