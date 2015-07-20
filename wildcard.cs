using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace WRPriorityReport {
  public class wildcard : Regex{
    public wildcard(string pattern):base(WildcardToRegex(pattern)){
     }
 
    public wildcard(string pattern, RegexOptions options):base(WildcardToRegex(pattern), options){
    }
 
    public static string WildcardToRegex(string pattern){
      return "^" + Regex.Escape(pattern).
       Replace("\\*", ".*").
       Replace("\\?", ".") + "$";
    }
  }
}
