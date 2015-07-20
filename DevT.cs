using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WRPriorityReport {
  public class DevT {
    public string WR;
    public string status;
    public string IV;
    public string IVstatus;
    public string NeedByDate;
    public string NeedByEvent;
    public DateTime createdDate;
    public string Project;
    public string IVseverity;

    public DevT(string WR_Num_As_String, string WR_Status, string InternalVariance, string InternalVarStatus, string dNBD, string dNBE, DateTime created, string PID, string IVSeverity) {
      WR = WR_Num_As_String;
      status = WR_Status;
      IV = InternalVariance;
      IVstatus = InternalVarStatus;
      if (dNBD == "") {
        NeedByDate = "Date not Set";
      } else if(dNBD.IndexOf(":")>0) {
        NeedByDate = dNBD.Substring(0, dNBD.IndexOf(":") - 3);
      }else{
        NeedByDate = dNBD;
      }
      NeedByEvent = dNBE;
      createdDate = created;
      switch (IVSeverity) {
        case "1":
          IVseverity = "Critical";
          break;
        case "2":
          IVseverity = "High";
          break;
        case "3":
          IVseverity = "Medium";
          break;
        case "4":
          IVseverity = "Low";
          break;
        case "5":
          IVseverity = "Enhancement";
          break;
        default:
          IVseverity = "Not set";
          break;
      }
       Project = PID;

    }
  }
}
