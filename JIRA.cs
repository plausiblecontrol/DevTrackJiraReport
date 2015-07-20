using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WRPriorityReport {
  public class WRQuery {
    public string JKey;
    public string Product;
    public string WR;
    public string Suite;
    public string CurrentRelease;
    public string Summary;
    public string IssueType;
    public string Projects;
    public string Description;
    public string Status;
    public string Resolution;
    public string RequestedDate;
    public string RequestedStatus;
    public string DueDate;
    public string Labels;
    public WRQuery(Jissues info, string[] validPIDs) {
      string temp = "";
      
      //JIRA is either a value, an object, or Null. TryCatches! *not proud*

      try { temp = unnull(info.key); } catch { }//JIRA ID    #A
      JKey = temp;//JIRA ID    #A
      temp = "";

      temp = "Nothing in JIRA";
      try { temp = unnull(info.fields.customfield_12802.value); } catch { }//Variances   #B
      Suite = temp;
      temp = "";
      try { temp = unnull(info.fields.customfield_10401); } catch { }//WR    #C
      WR = temp;
      temp = "";
      try { temp = unnull(info.fields.description); } catch { }//WR    #C
      Description = temp;
      temp = "";
      try { temp = String.Join(",", info.fields.labels); } catch { }//"CERT"
      Labels = temp;
      temp = "";
      try { temp = unnull(info.fields.status.name); } catch { }//WR    #C
      Status = temp;

      temp = "";
      try { temp = unnull(info.fields.resolution.name); } catch { temp = "Unresolved"; }//WR    #C
      Resolution = temp;

      temp = string.Join(Environment.NewLine, validPIDs);
      try { temp = filterOut(info.fields.customfield_10410, validPIDs); } catch { }//WR    #C
      Projects = temp;


      temp = "";
      try { temp = unnull(info.fields.issuetype.name); } catch { }//Issue Type    #D
      IssueType = temp;
      temp = "";
      try { temp = unnull(info.fields.customfield_10502.value); } catch { }//Product    #E
      Product = temp;
      temp = "";
      try { temp = unnull(info.fields.summary); } catch { }//Summary		#F
      Summary = temp;
      temp = "";
      temp = "";
      try { temp = unnull(info.fields.customfield_10800); } catch { }     //K
      CurrentRelease = temp;

      temp = "xxx";
      try { temp = unnull(info.fields.customfield_12100.value); } catch { }     //K
      RequestedStatus = temp;

      temp = "";
      try { temp = unnull(info.fields.customfield_12101); } catch { }     //K
      RequestedDate = temp;

      temp = "";
      try { temp = unnull(info.fields.duedate); } catch { }     //K
      DueDate = temp;

    }
    private static string unnull(object Value) {
      return Value == null ? "xxx" : Value.ToString();
    }
    private static string undate(string ugly) {
      //"2015-01-16T16:47:29.000-0600"

      //temp=11/26/2014 12:00:00 AM
      //releasedate=2014-11-26
      try {
        string[] a = ugly.Substring(0, ugly.IndexOf("T")).Split(new Char[] { '-' }).ToArray();
        string[] b = { a[1], a[2], a[0] };
        return string.Join("/", b);
      } catch {
        return ugly;
      }

    }
    private string filterOut(List<string> list, string[] keeps) {
      //string filter = "";
      List<string> betterKeeps = new List<string>();
      foreach (string keep in keeps) {
        if (keep.IndexOf(":") > 0) {
          betterKeeps.Add(keep.Substring(0, keep.IndexOf(":") - 1));
        } else if (keep.IndexOf(" ") > 0) {
          betterKeeps.Add(keep.Substring(0, keep.IndexOf(" ") - 1));
        } else {
          betterKeeps.Add(keep);
        }
      }
      foreach (string s in list) {
        foreach (string k in betterKeeps) {
          if (s.Contains(k)) {
            return s;
          }
        }
      }
      return "JIRA DNE";
    }
  }
}
