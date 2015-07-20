using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Net;
using System.IO;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web;
using System.Web.Script.Serialization;
using System.Drawing;

namespace WRPriorityReport {
  /// <summary>
  /// Interaction logic for MainWindow.xaml
  /// </summary>
  public partial class MainWindow : System.Windows.Window {
    public MainWindow() {
      
      
      DateTime newFile = File.GetLastWriteTime(@"\\WRPriorityReportTool\WRPriorityReport.exe");//update for proper location, used to check new version
      if (newFile > Convert.ToDateTime("7/1/15")) {
        MessageBoxResult dialogResult = MessageBox.Show("There is a new tool available!"+Environment.NewLine+"This version will not work, please get the newest.", "Please Update!", MessageBoxButton.OK);
        return;
      }
      
      
      string[] args = Environment.GetCommandLineArgs();
      List<string> fileInfo = new List<string>();
      if (args.Length > 0) {
        try {

          using (StreamReader sr = new StreamReader(args[1])) {
            String line;
            while ((line = sr.ReadLine()) != null) {
              fileInfo.Add(line);
            }
          }
        } catch {
          //bad file?
        }
      }

      InitializeComponent();
      if (fileInfo.Count > 0) {
        if (fileInfo[0] == "All Dates") {
          Dispatcher.Invoke((System.Action)delegate() {
            doAllcb.IsChecked = true;
            dtPick.SelectedDate = Convert.ToDateTime("2015/01/01");
          });
        } else {
          dtPick.SelectedDate = Convert.ToDateTime(fileInfo[0]);
        }
      } else {
        dtPick.SelectedDate = DateTime.Today.AddDays(-7);
      }
      ignoreBox.Text = "CAR/PAR/CI, Conversions and Builds, Customer Support, Hardware Products, Indigo: Internal, ITO Variances, Engineering Training, Product Planning, Systems Engineering Requests, AP1: In, APDCL1: In, KEPCO8: Kansas, NBSO1: New, NBSO1_U1: New, OECC6, WB1: In";
      BackgroundWorker bw1 = new BackgroundWorker();
      bw1.DoWork += new DoWorkEventHandler(getPList);
      bw1.RunWorkerAsync(fileInfo.ToArray());
      pBar.Visibility = Visibility.Hidden;
    }

    public void getPList(object sender, DoWorkEventArgs e) {
      List<string> rdr = getPID();
      string[] fileInfo = (string[])e.Argument;
      Dispatcher.Invoke((System.Action)delegate() {
        foreach (string pid in rdr) {
          customers.Items.Add(pid);
        }
        try {
          for (int i = 1; i < fileInfo.Length; i++) {
            customers.SelectedItems.Add(fileInfo[i]);
          }
        } catch {
          status.Content = "Error parsing the savefile.";
        }
        if (fileInfo.Length > 1) {
          List<string> projects = new List<string>();
          foreach (object selItem in customers.SelectedItems) {
            projects.Add(selItem.ToString());
          }
          BackgroundWorker bw1 = new BackgroundWorker();
          bw1.DoWork += new DoWorkEventHandler(verifyAndUpdate);
          bw1.RunWorkerAsync(projects);
        }

      });


    }
    public static string undate(string hasdate) {
      int yes = hasdate.IndexOf("12:00");//3/25/2015 12:00:00 AM because excel...
      int ffsJIRA = hasdate.IndexOf("-");//2015-03-25
      if (yes > 0) {
        return hasdate.Substring(0, yes - 1);
      } else if (ffsJIRA == 4 && hasdate.Length == 10) {
        //2015-03-25 to 3/25/2015
        string[] dd = hasdate.Split('-');
        string lest = String.Join("/", new[] { dd[1].TrimStart('0'), dd[2].TrimStart('0'), dd[0] });
        return lest;
      } else {
        return hasdate;
      }
    }
    public static int getSIndex(List<string> values1, List<string> values2, string value1, string value2) {
      int index = -1;
      for (int i = 0; i < values1.Count; i++) {
        if (values1[i] == value1 && values2[i] == value2) {
          index = i;
          break;
        }
      }
      return index;
    }
    public static string getDStatus(string WR, List<DevT> DevTrack) {
      foreach (DevT d in DevTrack) {
        if (d.WR == WR) {
          return d.status;
        }
      }
      return "N/A";
    }
    public static string getIVStatus(string WR, List<DevT> DevTrack) {
      foreach (DevT d in DevTrack) {
        if (d.WR == WR) {
          return d.IVstatus;
        }
      }
      return "N/A";
    }
    public static string getNBDE(string WR, List<DevT> DevTrack) {
      foreach (DevT d in DevTrack) {
        if (d.WR == WR) {
          string temp = d.NeedByEvent + Environment.NewLine + Environment.NewLine + d.NeedByDate;
          return temp;
        }
      }
      return "N/A";
    }
    public static string getNBE(string WR, List<DevT> DevTrack) {
      foreach (DevT d in DevTrack) {
        if (d.WR == WR) {
          return d.NeedByEvent;
        }
      }
      return "N/A";
    }
    public static string getNBD(string WR, List<DevT> DevTrack) {
      foreach (DevT d in DevTrack) {
        if (d.WR == WR) {
          return d.NeedByDate;
        }
      }
      return "N/A";
    }
    public static string getIV(string WR, List<DevT> DevTrack) {
      foreach (DevT d in DevTrack) {
        if (d.WR == WR) {
          return d.IV;
        }
      }
      return "N/A";
    }
    static List<string> getColumn(object[,] dataTable, int column) {
      List<string> data = new List<string>();
      int maxR = dataTable.GetLength(0);
      for (int i = 1; i <= maxR; i++) {
        try {
          data.Add(dataTable[i, column].ToString());
        } catch {
          data.Add("");
        }
      }
      return data;
    }
    public static List<string> getRow(object[,] dataTable, int row) {
      List<string> data = new List<string>();
      int maxC = dataTable.GetLength(1);
      for (int i = 1; i <= maxC; i++) {
        try {
          data.Add(dataTable[row, i].ToString());
        } catch {
          data.Add("");
        }
      }
      return data;
    }
    static int getIndex(List<string> values, string value) {
      int index = -1;
      for (int i = 0; i < values.Count; i++) {
        if (values[i] == value) {
          index = i;
          break;
        }
      }
      return index;
    }
    public static List<string> getNowClosed(List<DevT> DevTrack, List<string> OldWRnPID) {
      List<string> nc = new List<string>();
      List<string> newWRs = new List<string>();
      foreach (DevT NewDT in DevTrack) {
        newWRs.Add(NewDT.WR);
      }
      foreach (string OldWR in OldWRnPID) {
        int i = getIndex(newWRs, OldWR);
        //int i = getSIndex(newWRs, OldWRs, OldWR, Project);
        if (i < 0) {
          nc.Add(OldWR);
        }
      }
      return nc;
    }
    public static string getAlpha(int Range) { //theres gotta be a better way to do this..
      string val = "";
      switch (Range) {
        case 1:
          val = "A";
          break;
        case 2:
          val = "B";
          break;
        case 3:
          val = "C";
          break;
        case 4:
          val = "D";
          break;
        case 5:
          val = "E";
          break;
        case 6:
          val = "F";
          break;
        case 7:
          val = "G";
          break;
        case 8:
          val = "H";
          break;
        case 9:
          val = "I";
          break;
        case 10:
          val = "J";
          break;
        case 11:
          val = "K";
          break;
        case 12:
          val = "L";
          break;
        case 13:
          val = "M";
          break;
        case 14:
          val = "N";
          break;
        case 15:
          val = "O";
          break;
        case 16:
          val = "P";
          break;
        case 17:
          val = "Q";
          break;
        case 18:
          val = "R";
          break;
        case 19:
          val = "S";
          break;
        case 20:
          val = "T";
          break;
        case 21:
          val = "U";
          break;
        case 22:
          val = "V";
          break;
        case 23:
          val = "W";
          break;
        case 24:
          val = "X";
          break;
        case 25:
          val = "Y";
          break;
        case 26:
          val = "Z";
          break;
        case 27:
          val = "AA";
          break;
        case 28:
          val = "AB";
          break;
        default:
          val = "A";
          break;

      }
      return val;
    }
    public List<string> getPID() {
      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;
      bool fail = false;
      List<string> pids = new List<string>();
      try {
        con = new SqlConnection(String.Format("Data Source=sqlmn201\\{0}; Database={1}; Integrated Security=true", "PROD2000", "SWISEDB"));
        con.Open();
        string qryStr = "select ProjectName from Project where ProjectName not like '*%' and isactiveproject = 1 order by ProjectName";
        cmd = new SqlCommand(qryStr);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        while (rdr.Read()) {
          pids.Add(rdr[0].ToString());
        }
      } catch {
        rdr = null;
        fail = true;
      } finally {
        if (rdr != null) {
          rdr.Close();
        }
        if (con.State == ConnectionState.Open) {
          con.Close();
        }
      }
      Dispatcher.Invoke((System.Action)delegate() {
        if (fail) {
          status.Content = "Couldn't connect to DevTrack.";
        } else {
          status.Content = "";
        }
      });
      return pids;
    }

    public List<DevT> getWRplus(string PID) {
      SqlDataReader rdr = null;
      SqlConnection con = null;
      SqlCommand cmd = null;
      List<DevT> dis = new List<DevT>();
      try {
        con = new SqlConnection(String.Format("redacted"));//update to run
        con.Open();
        string qryStr = "";
        #region QueryString
//if old-style devtrack, else new
        if (PID.Contains("EVN1") || PID.Contains("LADWP1M_U6") || PID.Contains("KAMO Power") || PID.Contains("SRP18") || PID.Contains("SCS1: Internal") || PID.Contains("PWRCO11")) {
          qryStr = "Select WR, WRStatus, VAR, VARStatus, Left(VARNeedByDate, 10), Priority, WRDate, VARSeverity from ( " +
          "SELECT Bug.ProblemID as WR, Bug.DateCreated as WRDate, ProgressStatusTypes.ProgressStatusName as WRStatus, Bug2.ProblemID as VAR, ProgressStatusTypes2.ProgressStatusName as VARStatus, Custom.Desc_Custom_3 as VARNeedByDate, Custom.Desc_Custom_2 as Priority, Bug2.CrntBugTypeID as VARSeverity " +
          "FROM SWISEDB.dbo.Bug, SWISEDB.dbo.Bug Bug2, SWISEDB.dbo.CustomerFieldTrackExt Custom, SWISEDB.dbo.BugLinks, SWISEDB.dbo.Project, SWISEDB.dbo.ProgressStatusTypes, SWISEDB.dbo.ProgressStatusTypes ProgressStatusTypes2 " +
          "WHERE " +
          "Bug.ProjectID = 17 and Bug2.ProjectID = Project.ProjectID and " +
          "BugLinks.LinkedBugID = Bug.BugID and " +
          "BugLinks.LinkedProjectID = Bug.ProjectID and " +
          "BugLinks.ProjectID = Project.ProjectID and " +
          "Bug2.BugID = BugLinks.BugID and " +
          "Bug2.ProjectID = BugLinks.ProjectID and " +
          "Bug.ProgressStatusID = ProgressStatusTypes.ProgressStatusID and " +
          "Bug.ProjectID = ProgressStatusTypes.ProjectID and " +
          "Bug2.ProgressStatusID = ProgressStatusTypes2.ProgressStatusID and " +
          "Bug2.ProjectID = ProgressStatusTypes2.ProjectID and " +
          "Bug2.ProjectID = Custom.ProjectID and " +
          "Bug2.BugID = Custom.BugID and " +
          "Project.ProjectName like '" + PID + "%' " +
          "UNION " +
          "SELECT Bug.ProblemID as WR, Bug.DateCreated as WRDate, ProgressStatusTypes.ProgressStatusName as WRStatus, Bug2.ProblemID as VAR, ProgressStatusTypes2.ProgressStatusName as VARStatus, Custom.Desc_Custom_3 as VARNeedByDate, Custom.Desc_Custom_2 as Priority, Bug2.CrntBugTypeID as VARSeverity " +
          "FROM SWISEDB.dbo.Bug, SWISEDB.dbo.Bug Bug2, SWISEDB.dbo.CustomerFieldTrackExt Custom, SWISEDB.dbo.BugLinks, SWISEDB.dbo.Project, SWISEDB.dbo.ProgressStatusTypes, SWISEDB.dbo.ProgressStatusTypes ProgressStatusTypes2 " +
          "WHERE " +
          "Bug.ProjectID = 17 and Bug2.ProjectID = Project.ProjectID and " +
          "BugLinks.BugID = Bug.BugID and " +
          "BugLinks.ProjectID = Bug.ProjectID and " +
          "BugLinks.LinkedProjectID = Project.ProjectID and " +
          "Bug2.BugID = BugLinks.LinkedBugID and " +
          "Bug2.ProjectID = BugLinks.LinkedProjectID and " +
          "Bug.ProgressStatusID = ProgressStatusTypes.ProgressStatusID and " +
          "Bug.ProjectID = ProgressStatusTypes.ProjectID and " +
          "Bug2.ProgressStatusID = ProgressStatusTypes2.ProgressStatusID and " +
          "Bug2.ProjectID = ProgressStatusTypes2.ProjectID and " +
          "Bug2.ProjectID = Custom.ProjectID and " +
          "Bug2.BugID = Custom.BugID and " +
          "Project.ProjectName like '" + PID + "%' ) q ";
        } else {
          qryStr = "Select WR, WRStatus, VAR, VARStatus, VARNeedByDate, ForwardTypes.ForwardTypeName as Priority, WRDate, VARSeverity from ( " +
            "SELECT Bug.ProblemID as WR, Bug.DateCreated as WRDate, ProgressStatusTypes.ProgressStatusName as WRStatus, Bug2.ProblemID as VAR, ProgressStatusTypes2.ProgressStatusName as VARStatus, Bug2.TaskPlannedStartDate as VARNeedByDate, Bug2.CrntForwardTypeID as VARPriority, Bug2.CrntBugTypeID as VARSeverity " +
            "FROM SWISEDB.dbo.Bug, SWISEDB.dbo.Bug Bug2, SWISEDB.dbo.BugLinks, SWISEDB.dbo.Project, SWISEDB.dbo.ProgressStatusTypes, SWISEDB.dbo.ProgressStatusTypes ProgressStatusTypes2 " +
            "WHERE " +
            "Bug.ProjectID = 17 and Bug2.ProjectID = Project.ProjectID and " +
            "BugLinks.LinkedBugID = Bug.BugID and " +
            "BugLinks.LinkedProjectID = Bug.ProjectID and " +
            "BugLinks.ProjectID = Project.ProjectID and " +
            "Bug2.BugID = BugLinks.BugID and " +
            "Bug2.ProjectID = BugLinks.ProjectID and " +
            "Bug.ProgressStatusID = ProgressStatusTypes.ProgressStatusID and " +
            "Bug.ProjectID = ProgressStatusTypes.ProjectID and " +
            "Bug2.ProgressStatusID = ProgressStatusTypes2.ProgressStatusID and " +
            "Bug2.ProjectID = ProgressStatusTypes2.ProjectID and " +
            "Project.ProjectName like '" + PID + "%' " +
            "UNION " +
            "SELECT Bug.ProblemID as WR, Bug.DateCreated as WRDate, ProgressStatusTypes.ProgressStatusName as WRStatus, Bug2.ProblemID as VAR, ProgressStatusTypes2.ProgressStatusName as VARStatus, Bug2.TaskPlannedStartDate as VARNeedByDate, Bug2.CrntForwardTypeID as VARPriority, Bug2.CrntBugTypeID as VARSeverity " +
            "FROM SWISEDB.dbo.Bug, SWISEDB.dbo.Bug Bug2, SWISEDB.dbo.BugLinks, SWISEDB.dbo.Project, SWISEDB.dbo.ProgressStatusTypes, SWISEDB.dbo.ProgressStatusTypes ProgressStatusTypes2 " +
            "WHERE " +
            "Bug.ProjectID = 17 and Bug2.ProjectID = Project.ProjectID and " +
            "BugLinks.BugID = Bug.BugID and " +
            "BugLinks.ProjectID = Bug.ProjectID and " +
            "BugLinks.LinkedProjectID = Project.ProjectID and " +
            "Bug2.BugID = BugLinks.LinkedBugID and " +
            "Bug2.ProjectID = BugLinks.LinkedProjectID and " +
            "Bug.ProgressStatusID = ProgressStatusTypes.ProgressStatusID and " +
            "Bug.ProjectID = ProgressStatusTypes.ProjectID and " +
            "Bug2.ProgressStatusID = ProgressStatusTypes2.ProgressStatusID and " +
            "Bug2.ProjectID = ProgressStatusTypes2.ProjectID and " +
            "Project.ProjectName like '" + PID + "%' ) q " +
            "Left Join ForwardTypes on ForwardTypes.ProjectID = 1450 and ForwardTypes.OrderNo+1 = VARPriority";
        }

        #endregion

        cmd = new SqlCommand(qryStr);
        cmd.Connection = con;
        rdr = cmd.ExecuteReader();
        while (rdr.Read()) {
          string dWR = "";
          string dWRstatus = "";
          string dIVnum = "";
          string dIVstatus = "";
          string dNBD = "Date Not Set";//need by date
          string dNBE = "Event Not Set";//need by event
          string dateString = "";
          string IVseverity = "";
          DateTime created = new DateTime();
          //rdr[6] date as 7/29/2014 11:17:13 AM
          try {
            dWR = rdr[0].ToString().Substring(3, rdr[0].ToString().Length - 3);
          } catch { }
          try { dWRstatus = rdr[1].ToString(); } catch { }
          if (dWR == "70996") {
            //do nothing
          }
          try { dIVnum = rdr[2].ToString(); } catch { }
          try { dIVstatus = rdr[3].ToString(); } catch { }
          try {
            dNBD = rdr[4].ToString();
            if (dNBD == "" || dNBD == "12/31/1969 6:00:00 PM")
              dNBD = "Date Not Set";
          } catch { dNBD = "Date Not Set"; }
          try {
            dNBE = rdr[5].ToString();
            if (dNBE == "")
              dNBE = "Event Not Set";
          } catch { dNBE = "Event Not Set"; }
          try { dateString = rdr[6].ToString(); } catch { dateString = "Unknown"; }
          try { IVseverity = rdr[7].ToString(); } catch { IVseverity = "Unlisted"; }
          created = DateTime.Parse(dateString);
          DevT d = new DevT(dWR, dWRstatus, dIVnum, dIVstatus, dNBD, dNBE, created, PID, IVseverity);
          dis.Add(d);
        }
        rdr.NextResult();
      } catch (Exception e) {
        string m = e.Message;
        rdr = null;
      } finally {
        if (rdr != null) {
          rdr.Close();
        }
        if (con.State == ConnectionState.Open) {
          con.Close();
        }
      }
      return dis;
    }

    private void createBtn_Click(object sender, RoutedEventArgs e) {
      List<string> items = new List<string>();
      if (ProjectsStr.Text == "") {
        foreach (var item in customers.SelectedItems) {
          items.Add(item.ToString());
        }
      } else {
        items = customers.Items.OfType<string>().ToList();
      }
      if (items.Count != 0) {
        BackgroundWorker bw1 = new BackgroundWorker();
        bw1.DoWork += new DoWorkEventHandler(verifyAndCommit);
        bw1.RunWorkerAsync(items);
      } else {
        MessageBoxResult dialogResult = MessageBox.Show("Please select a project", "WR Priority Report", MessageBoxButton.OK);
      }
    }

    private void verifyAndCommit(object sender, DoWorkEventArgs e) {
      List<string> customerList = (List<string>)e.Argument;
      List<string> updatedList = new List<string>();
      string projectQry = "";
      disableButtons(true);
      Dispatcher.Invoke((System.Action)delegate() {
        projectQry = ProjectsStr.Text;
        updatedList = customers.SelectedItems.OfType<string>().ToList();
      });
      if (projectQry != "") {
        string[] pQOs = projectQry.Split(',');
        foreach (string projectQ in pQOs) {
          wildcard chaos = new wildcard(projectQ, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
          foreach (string projectCode in customerList) {
            if (chaos.IsMatch(projectCode))
              if (getIndex(updatedList, projectCode) < 0) {
                updatedList.Add(projectCode);
              }
          }
        }
      } else {
        updatedList = customerList;
      }

      string[] pid = updatedList.ToArray();
      MessageBoxResult dialogResult = new MessageBoxResult();
      if (pid.Length > 0 && pid.Length < 11) {
        dialogResult = MessageBox.Show("This will generate a report for: " + Environment.NewLine + string.Join(" & " + Environment.NewLine, pid), "WR Priority Report", MessageBoxButton.OKCancel);
      } else if (pid.Length >= 11) {
        dialogResult = MessageBox.Show("This will generate a report for " + pid.Length + " projects.", "WR Priority Report", MessageBoxButton.OKCancel);
      } else {
        dialogResult = MessageBox.Show("No projects found to match", "WR Priority Report", MessageBoxButton.OK);
      }
      if (dialogResult == MessageBoxResult.OK && pid.Length > 0) {
        WRPR(pid);

      } else {
        disableButtons(false);
      }
    }

    private void Button_Click(object sender, RoutedEventArgs e) {
      Dispatcher.Invoke((System.Action)delegate() {
        customers.UnselectAll();
      });
    }

    public List<string> makeUrl(List<DevT> wrList) {//labels = CERT
      List<string> urls = new List<string>();
      string url = "";
      int count = wrList.Count;
      for (int i = 0; i <= count - 1; i++) {
        if (wrList[i] != null) {
          if (url.Length != 0) {
            url += "%20OR%20";
          }
          url += "WR%20~%20\"" + wrList[i].WR + "\"";
          if (i % 30 == 0 && i != 0) {
            urls.Add(url);
            url = "";
          }
        }
      }
      if (url != "") {
        urls.Add(url);
      }
      return urls;
    }

    public List<string> mkoldURLs(List<string> OldWRs) {
      List<string> urls = new List<string>();
      string url = "";
      int count = OldWRs.Count;
      for (int i = 0; i <= count - 1; i++) {
        if (OldWRs[i] != null) {
          if (url.Length != 0) {
            url += "%20OR%20";
          }
          url += "WR%20~%20\"" + OldWRs[i] + "\"";
          if (i % 30 == 0 && i != 0) {
            urls.Add(url);
            url = "";
          }
        }
      }
      if (url != "") {
        urls.Add(url);
      }
      return urls;
    }

    public int DevWRExists(List<WRQuery> Issues, string WR) {
      int index = -1;
      for (int i = 0; i < Issues.Count; i++) {
        if (Issues[i].WR == WR) {
          index = i;
          break;
        }
      }

      return index;
    }

    public WR pJSON(string qTxt) {//list<string> qTxts

      JiraResource resource = new JiraResource();
      string bURL = "https://redacted.atlassian.net/rest/api/latest/search?jql="; //base url for JIRA access, update to run
      string JSONdata = null;
      int statusCode = 0;
      Stream s;
      StreamReader r;
      HttpWebResponse webRes;
      HttpWebRequest WebReq = WebRequest.Create(bURL + qTxt) as HttpWebRequest;
      //WebReq.Timeout = Timeout.Infinite;//solved JIRA lagginess, otherwise Error will report 'Error accessing JIRA'
      //WebReq.KeepAlive = true;//also was required
      WebReq.ContentType = "application/json";
      WebReq.Method = "GET";
      WebReq.Headers["Authorization"] = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes(resource.m_Username + ":" + resource.m_Password));

      try {
        webRes = (HttpWebResponse)WebReq.GetResponse();
        s = webRes.GetResponseStream();
        r = new StreamReader(s);
        JSONdata = r.ReadToEnd();
        statusCode = (int)webRes.StatusCode;
        s.Close();
        r.Close();
      } catch (WebException e) {
        s = e.Response.GetResponseStream();
        r = new StreamReader(s);
        JSONdata = r.ReadToEnd();
        statusCode = (int)((HttpWebResponse)e.Response).StatusCode;
        s.Close();
        r.Close();
      }
      try {
        WR tem = new JavaScriptSerializer().Deserialize<WR>(JSONdata);
        return tem;
      } catch (Exception e) {
        string m = e.Message;
        MessageBoxResult dialogResult = new MessageBoxResult();
        dialogResult = MessageBox.Show("ERROR:" + m, "Error accessing JIRA", MessageBoxButton.OK);

      }
      return null;
    }

    public static string RemoveSpecialCharacters(string str) {
      StringBuilder sb = new StringBuilder();
      foreach (char c in str) {
        if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '.' || c == '_') {
          sb.Append(c);
        }
      }
      return sb.ToString();
    }

    static void releaseObject(object obj) {
      try {
        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        obj = null;
      } catch (Exception ex) {
        obj = null;
        Console.WriteLine("Exception Occured while releasing object " + ex.ToString());
      } finally {
        GC.Collect();
        GC.WaitForPendingFinalizers();
      }
    }

    private List<DevT> OpenDevT(string[] pids, string[] ignoreList) {
      List<DevT> DevTrackList = new List<DevT>();
      List<DevT> OpenDevItems = new List<DevT>();
      double j = 100.0;
      double k = j / pids.Length;

      /*slow but safe*/
      //foreach (string pid in pids) {
      //  DevTrackList.AddRange(getWRplus(pid));
      //  Dispatcher.Invoke((System.Action)delegate() {
      //    pBar.Value += k;
      //  });
      //  Thread.Sleep(111);
      //}
      /**/
      //List<string> sPIDs = pids.ToList();
      for (int n = 0; n < pids.Count(); n++) {
        for (int m = 0; m < ignoreList.Count(); m++) {
          if (pids[n].Contains(ignoreList[m])) {
            pids[n] = "asdf_NAN_asdf";
          }
        }
      }

      /*fast but reckless*/
      Parallel.ForEach(pids, pid => {
        DevTrackList.AddRange(getWRplus(pid));
        Dispatcher.Invoke((System.Action)delegate() {
          pBar.Value += k;
        });
      });
      Dispatcher.Invoke((System.Action)delegate() {
        pBar.Value += k;
      });
      /**/




      Dispatcher.Invoke((System.Action)delegate() {
        status.Content = "Checking for errors..." + Environment.NewLine + "Shouldn't take longer than 15sec.";
      });
      string[] debug = new string[3];
      int dLi = 0;
      try {
        foreach (DevT DTitem in DevTrackList) {
          if (DTitem.status != "Confirm Verified" && DTitem.status != "Confirm Duplicate" && DTitem.status != "Confirm Reject" && DTitem.IVstatus != "Closed") {
            OpenDevItems.Add(DTitem);
          }
          dLi++;
          debug[0] = dLi.ToString();
          debug[1] = DTitem.Project;
          debug[2] = DTitem.WR;
        }
      } catch (Exception e) {
        MessageBoxResult dialogResult = new MessageBoxResult();
        dialogResult = MessageBox.Show("Investigate " + debug[1] + " after WR" + debug[2] + Environment.NewLine + debug[0] + "ERROR:" + e.Message, "Fatal error with DevTrack - TryAgain?", MessageBoxButton.OK);

      }

      return OpenDevItems;
    }

    private void WRPR(string[] pids) {
      string dt = DateTime.Now.Day.ToString("00") + DateTime.Now.Month.ToString("00") + DateTime.Now.Year.ToString().Substring(2, 2);
      Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
      dlg.FileName = "_" + dt + "_WRReports";
      dlg.DefaultExt = ".xlsx";
      dlg.Filter = "Excel Workbook (.xlsx)|*.xlsx";
      string[] debug = new string[4];
      Nullable<bool> result = dlg.ShowDialog();
      List<DevT> DevTrackList = new List<DevT>();
      if (result == true) {
        string textSave = dlg.FileName.Substring(0, dlg.FileName.Length - 4) + "txt";
        string dtPicked = "";//2015-01-27
        DateTime selectedDT = new DateTime();
        bool doAll = false;
        Dispatcher.Invoke((System.Action)delegate() {
          selectedDT = dtPick.SelectedDate.Value;
          string dtDay = dtPick.SelectedDate.Value.Day.ToString("00");
          string dtMonth = dtPick.SelectedDate.Value.Month.ToString("00");
          string dtYear = dtPick.SelectedDate.Value.Year.ToString("0000");
          dtPicked = dtYear + "-" + dtMonth + "-" + dtDay;
          doAll = doAllcb.IsChecked.Value;
        });
        using (StreamWriter sw = File.CreateText(textSave)) {
          if (doAll) {
            sw.WriteLine("All Dates");
          } else {
            sw.WriteLine(dtPicked);
          }
          foreach (string pid in pids) {
            sw.WriteLine(pid);
          }
        }
        string ignoreItem = "qwertyuiop12345";
        Dispatcher.Invoke((System.Action)delegate() {
          status.Content = "Retrieving data from DevTrack...";
          pBar.Visibility = Visibility.Visible;
          pBar.Value = 0;
          ignoreItem = ignoreBox.Text;
        });

        //ignorebox
        string[] ignoreItems = System.Text.RegularExpressions.Regex.Split(ignoreItem, ", ");

        List<DevT> OpenDevItems = OpenDevT(pids, ignoreItems);
        List<DevT> LatestOpenDevItems = new List<DevT>();
        List<WR> JIRAitems = new List<WR>();
        bool isError = false;

        foreach (DevT DTitem in OpenDevItems) {
          if (DTitem == null)
            isError = true;
        }
        if (isError) {
          Dispatcher.Invoke((System.Action)delegate() {
            pBar.Visibility = System.Windows.Visibility.Hidden;
            status.Content = "DevTrack possibly busy/unavailable, trying to reconnect." + Environment.NewLine + "Please wait, retries occur ~3seconds..";
            pBar.Value = 0;
          });
          while (isError) {
            Thread.Sleep(3000);
            OpenDevItems = OpenDevT(pids, ignoreItems);
            for (int i = 0; i < OpenDevItems.Count; i++) {
              if (OpenDevItems[i] == null) {
                isError = true;
                break;
              } else {
                isError = false;
              }
            }
          }
        }



        foreach (DevT DTitem in OpenDevItems) {
          if (DTitem == null)
            isError = true;
        }
        if (!isError) {
          #region meatNpotatoes
          if (!doAll) {

            foreach (DevT DTitem in OpenDevItems) {
              if (DTitem.createdDate >= selectedDT) {
                LatestOpenDevItems.Add(DTitem);
              }
            }
          } else {
            LatestOpenDevItems = OpenDevItems;
          }
          List<string> urls = makeUrl(LatestOpenDevItems);
          double j = 100.0;
          double k = j / urls.Count();

          Dispatcher.Invoke((System.Action)delegate() {
            status.Content = "Querying a lot of JIRA...";
            pBar.Visibility = Visibility.Visible;
            pBar.Value = 0;
          });

          //int speed = 987;//get wait

          //foreach (string qry in urls) {//parallel could be a LOT faster, but might ddos jira if its lagging.
          //  JIRAitems.Add(pJSON(qry));
          //  Dispatcher.Invoke((System.Action)delegate() {
          //    pBar.Value += k;
          //  });
          //  Thread.Sleep(speed);
          //}

          Parallel.ForEach(urls, qry => {
            JIRAitems.Add(pJSON(qry));
            Dispatcher.Invoke((System.Action)delegate() {
              pBar.Value += k;
            });
          });

          Dispatcher.Invoke((System.Action)delegate() {
            status.Content = "Parsing JIRA data...";
          });
          List<WRQuery> JItemsC = new List<WRQuery>();
          foreach (WR dd in JIRAitems) {
            foreach (Jissues info in dd.issues) {
              JItemsC.Add(new WRQuery(info, pids));
            }
          }

          Dispatcher.Invoke((System.Action)delegate() {
            status.Content = "Finding JIRA errors...";
          });
          List<string> notInJIRA = new List<string>();
          foreach (DevT lodi in LatestOpenDevItems) {
            if (DevWRExists(JItemsC, lodi.WR) < 0) {
              notInJIRA.Add(lodi.WR);
            }
          }

          Dispatcher.Invoke((System.Action)delegate() {
            status.Content = "No WRs were found in criteria!";
            pBar.Value = 0;
          });
          bool worked = false;
          if (JIRAitems.Count != 0) {
            Dispatcher.Invoke((System.Action)delegate() {
              status.Content = "Wrestling Excel...";
              pBar.Value = 0;
            });



            //EXCEL should be on its own to better manage the crashes
            //can't pass excel objects for some reason, investigate to clean up code redundancy
            List<string> cautions = new List<string>();
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;
            Excel.Worksheet dataSheet = null;
            Excel.Range xlR = null;
            try {
              excelApp = new Excel.Application();
              workbook = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
              sheets = workbook.Sheets;
              dataSheet = sheets[1];
              string dtNow = DateTime.Now.Month.ToString("00") + "-" + DateTime.Now.Day.ToString("00") + "-" + DateTime.Now.Year.ToString().Substring(2, 2);
              if (!doAll) {
                dataSheet.Name = dtPicked + " and newer";
              } else {
                dataSheet.Name = "All";
              }

              #region Header format
              //string[] headers = { "PD Priority Submission Date", "Updated?" + Environment.NewLine + "(Yes/No)", "Project", "DevTeam", "WR#", "WR Status (JIRA -" + Environment.NewLine + "Resolution/Status)", "WR Created" + Environment.NewLine + "Date", "WR Summary", "WR Description", "WR Product", "Product" + Environment.NewLine + "Version", "Project's Variance", "Variance Severity", "Variance's Need by" + Environment.NewLine + "Event (Var Priority)", "Variance's Need by" + Environment.NewLine + "Date", "PE/PSE Mgmt Priority" + Environment.NewLine + "(High/Med/Low)", "PE/PSE Mgmt Justification", "Certification Priority" + Environment.NewLine + "(High/Med/Low)", "Certification Justification", "Project Delivery Priority" + Environment.NewLine + "(High/Med/Low)", "Project Delivery Need by Date","WR Count"+Environment.NewLine+"(Duplicates)" };
              string[] headers = { "PD Priority" + Environment.NewLine + "Submission Date", "Updated?" + Environment.NewLine + "(Yes/No)", "Project", "DevTeam", "WR#", "WR Resolution/" + Environment.NewLine + "Status", "Request" + Environment.NewLine + "Status", "Requested" + Environment.NewLine + "Date", "Planned" + Environment.NewLine + "Release Date", "WR Created" + Environment.NewLine + "Date", "WR Summary", "WR Description", "WR Product", "Product" + Environment.NewLine + "Version", "Project's Variance", "Variance Severity", "Variance's Need by" + Environment.NewLine + "Event (Var Priority)", "Variance's Need by" + Environment.NewLine + "Date", "PE/PSE Mgmt Priority" + Environment.NewLine + "(Urgent/High/Med/Low)", "PE/PSE Mgmt Justification", "Certification Priority" + Environment.NewLine + "(Urgent/High/Med/Low)", "Certification Justification", "Project Delivery Priority" + Environment.NewLine + "(Urgent/High/Med/Low)", "Project Delivery Need by Date", "WR Count" + Environment.NewLine + "(Duplicates)", "QC" + Environment.NewLine + "(PE/PSE Justification)", "QC" + Environment.NewLine + "(Var. Date)" };
              object[,] xlHeader = new object[1, headers.Length];
              string columnMax = getAlpha(headers.Length);
              for (int i = 0; i < headers.Length; i++) {
                xlHeader[0, i] = headers[i];
              }
              xlR = dataSheet.Range["A1", columnMax + "1"];//Header
              xlR.Value2 = xlHeader;
              #endregion

              dataSheet.get_Range("G:G", Type.Missing).EntireColumn.ColumnWidth = 50;
              int nXLcount = LatestOpenDevItems.Count();
              double iB = 100.0;
              double iJ = iB / nXLcount;
              int rowCount = 2; //start at row 2

              foreach (DevT DT in LatestOpenDevItems) {
                debug[0] = DT.WR;
                debug[1] = DT.Project;
                debug[3] = rowCount.ToString();
                int xWR = DevWRExists(JItemsC, DT.WR);
                string wSuite = "WR NOT IN JIRA!";
                string wSummary = "WR NOT IN JIRA!";
                string wCurRel = "WR NOT IN JIRA!";
                string wIT = "WR NOT IN JIRA!";
                string wProd = "WR NOT IN JIRA!";
                string wStatus = "WR NOT IN JIRA!";
                string rStatus = "WR NOT IN JIRA!";
                string rDate = "WR NOT IN JIRA!";
                string dDate = "WR NOT IN JIRA!";
                string wDesc = "WR NOT IN JIRA! Contact Development for assistance with this WR.";
                if (xWR >= 0) {
                  wSuite = JItemsC[xWR].Suite;
                  rStatus = JItemsC[xWR].RequestedStatus;
                  rDate = JItemsC[xWR].RequestedDate;
                  dDate = JItemsC[xWR].DueDate;
                  wSummary = JItemsC[xWR].Summary;
                  wCurRel = JItemsC[xWR].CurrentRelease;
                  wIT = JItemsC[xWR].IssueType;
                  wStatus = JItemsC[xWR].Resolution + Environment.NewLine + JItemsC[xWR].Status;
                  wProd = JItemsC[xWR].Product;
                  wDesc = JItemsC[xWR].Description;
                } string cdate = DT.createdDate.Month + "-" + DT.createdDate.Day + "-" + DT.createdDate.Year;
                string[] newLine = { dtNow, "No", DT.Project, wSuite, DT.WR, wStatus, rStatus, rDate, dDate, cdate, wSummary, "'" + wDesc, wProd, wCurRel, DT.IV, DT.IVseverity, DT.NeedByEvent, DT.NeedByDate };
                //string[] newLine = { dtNow, "No", DT.Project, wSuite, DT.WR, wStatus, rStatus, rDate, dDate, cdate, wSummary, wDesc, wProd, wCurRel, DT.IV, DT.IVseverity, DT.NeedByEvent, DT.NeedByDate };//DEBUG ONLY
                object[,] xlNewLine = new object[1, newLine.Length];
                for (int i = 0; i < newLine.Length; i++) {
                  xlNewLine[0, i] = newLine[i];
                }
                xlR = dataSheet.Range["A" + rowCount, getAlpha(newLine.Length) + rowCount];
                xlR.Value2 = xlNewLine;
                rowCount++;
                Dispatcher.Invoke((System.Action)delegate() {
                  pBar.Value += iJ;
                });
              }//end of magic

              #region findDuplicates
              int maxRf = dataSheet.UsedRange.Rows.Count;
              object[,] arr = dataSheet.get_Range("C2:E" + maxRf).Value;
              List<string> ProjectRow = new List<string>();
              List<string> ProjectWR = new List<string>();
              for (int i = 1; i < maxRf; i++) {
                ProjectRow.Add(arr[i, 1].ToString());
                ProjectWR.Add(arr[i, 3].ToString());
              }
              List<string> duplicatedWRs = ProjectWR.GroupBy(x => x)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key).ToList();

              if (duplicatedWRs.Count > 0) {
                Excel.Worksheet dpSheet = (Worksheet)sheets.Add(Type.Missing, sheets[1], Type.Missing, Type.Missing);
                dpSheet.Name = "Duplicated WRs";
                int rowToPrint = 2;
                dpSheet.Cells[1, 1] = "Duplicated" + Environment.NewLine + "WRs Found";
                dpSheet.Cells[1, 2] = "Projects found linked";
                foreach (string dWRx in duplicatedWRs) {
                  debug[3] = rowToPrint.ToString();
                  string ProjectsFound = "";
                  for (int i = 0; i < ProjectWR.Count; i++) {
                    if (dWRx == ProjectWR[i]) {
                      if (ProjectsFound == "") {
                        ProjectsFound += ProjectRow[i];
                      } else {
                        ProjectsFound += Environment.NewLine + ProjectRow[i];
                      }
                    }
                  }
                  //print this WR
                  dpSheet.Cells[rowToPrint, 1] = dWRx;
                  dpSheet.Cells[rowToPrint, 2] = ProjectsFound;
                  rowToPrint++;
                }
                dpSheet.get_Range("A:A", Type.Missing).EntireColumn.ColumnWidth = 16;
                dpSheet.get_Range("B:B", Type.Missing).EntireColumn.ColumnWidth = 32;
                dpSheet.get_Range("A:B", Type.Missing).EntireColumn.VerticalAlignment = XlVAlign.xlVAlignCenter;
                dpSheet.get_Range("A:B", Type.Missing).EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                dpSheet.get_Range("A1", "B1").Cells.Font.Bold = true;
                dpSheet.Select();
                releaseObject(dpSheet);
              }


              #endregion
              dataSheet.get_Range("A:" + columnMax, Type.Missing).EntireColumn.VerticalAlignment = XlVAlign.xlVAlignCenter;
              dataSheet.get_Range("A:F", Type.Missing).EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;
              dataSheet.get_Range("A:" + columnMax, Type.Missing).EntireColumn.WrapText = true;
              dataSheet.get_Range("A:" + columnMax, Type.Missing).EntireColumn.ColumnWidth = 17.8;
              dataSheet.get_Range("B:B", Type.Missing).EntireColumn.ColumnWidth = 12;
              dataSheet.get_Range("E:E", Type.Missing).EntireColumn.ColumnWidth = 8;
              dataSheet.get_Range("G:G", Type.Missing).EntireColumn.ColumnWidth = 14;
              dataSheet.get_Range("H:J", Type.Missing).EntireColumn.ColumnWidth = 12;
              dataSheet.get_Range("K:K", Type.Missing).EntireColumn.ColumnWidth = 32;
              dataSheet.get_Range("L:L", Type.Missing).EntireColumn.ColumnWidth = 50;
              dataSheet.get_Range("Q:R", Type.Missing).EntireColumn.ColumnWidth = 20;
              dataSheet.get_Range("S:X", Type.Missing).EntireColumn.ColumnWidth = 28;
              dataSheet.get_Range("A:" + columnMax, Type.Missing).EntireRow.HorizontalAlignment = XlHAlign.xlHAlignLeft;
              //Column data validation
              string validS = "Urgent,High,Med,Low,SCT";
              string validW = "Urgent,High,Med,Low";
              //string validC = "High,Med,Low,Not required";
              string validU = "Yes,No";
              dataSheet.get_Range("B2", "B" + (rowCount - 1)).Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, validU, Type.Missing);
              dataSheet.get_Range("S2", "S" + (rowCount - 1)).Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, validS, Type.Missing);
              dataSheet.get_Range("W2", "W" + (rowCount - 1)).Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, validW, Type.Missing);
              dataSheet.get_Range("U2", "U" + (rowCount - 1)).Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, validW, Type.Missing);

              xlR = dataSheet.get_Range("A1:" + columnMax + (rowCount - 1));
              dataSheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, xlR, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = "MyTableStyle";
              dataSheet.ListObjects.get_Item("MyTableStyle").TableStyle = "TableStyleLight8";
              dataSheet.get_Range("A1", columnMax + "1").EntireRow.RowHeight = 31;
              dataSheet.get_Range("A1", "B1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Goldenrod);
              dataSheet.get_Range("R1", "R1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Brown);
              dataSheet.get_Range("S1", "X1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Goldenrod);
              dataSheet.get_Range("Y1", "AA1").Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DeepSkyBlue);
              dataSheet.get_Range("Y:" + columnMax, Type.Missing).EntireColumn.Hidden = true;
              //dataSheet.get_Range("G:I", Type.Missing).EntireColumn.Hidden = true;
              dataSheet.Range["Y2"].Formula = "=COUNTIF(E:E,[@[WR'#]])";//duplicate formula
              dataSheet.Range["X2"].Formula = "=[@[Variance''s Need by" + Environment.NewLine + "Date]]";
              dataSheet.Columns[24].NumberFormat = "m/d/yyyy";

              dataSheet.Range["W2"].Value = "=IF([@[PE/PSE Mgmt Priority" + Environment.NewLine + //gotta use .Value because of length I guess?
                "(Urgent/High/Med/Low)]]<>\"SCT\",[@[PE/PSE Mgmt Priority" + Environment.NewLine +
                  "(Urgent/High/Med/Low)]],[@[Certification Priority" + Environment.NewLine +
                  "(Urgent/High/Med/Low)]])";
              dataSheet.Range["Z2"].Value = "=IF(AND(OR([@[PE/PSE Mgmt Priority" + Environment.NewLine +
                "(Urgent/High/Med/Low)]]=\"Urgent\",[@[PE/PSE Mgmt Priority" + Environment.NewLine +
                "(Urgent/High/Med/Low)]]=\"High\",[@[PE/PSE Mgmt Priority" + Environment.NewLine +
                "(Urgent/High/Med/Low)]]=\"Med\"),[@[PE/PSE Mgmt Justification]]=\"\"),\"Not Justified!\",\"Justified\")";
              dataSheet.Range["AA2"].Value = "=IF(AND(OR([@[PE/PSE Mgmt Priority" + Environment.NewLine +
                "(Urgent/High/Med/Low)]]=\"Urgent\",[@[PE/PSE Mgmt Priority" + Environment.NewLine +
                "(Urgent/High/Med/Low)]]=\"High\",[@[PE/PSE Mgmt Priority" + Environment.NewLine +
                "(Urgent/High/Med/Low)]]=\"Med\"),[@[Variance''s Need by" + Environment.NewLine +
                "Date]]=\"Date Not Set\"),\"Req'd Date Not Set!\",IF(AND(OR([@[PE/PSE Mgmt Priority" + Environment.NewLine +
                "(Urgent/High/Med/Low)]]=\"Urgent\",[@[PE/PSE Mgmt Priority" + Environment.NewLine +
                "(Urgent/High/Med/Low)]]=\"High\",[@[PE/PSE Mgmt Priority" + Environment.NewLine +
                "(Urgent/High/Med/Low)]]=\"Med\"),[@[Variance''s Need by" + Environment.NewLine +
                "Date]]<=TODAY()),\"Old Date Set!\",\"Adequate Date\"))";


              ((Excel._Worksheet)dataSheet).Activate();
              dataSheet.Application.ActiveWindow.SplitRow = 1;
              dataSheet.Application.ActiveWindow.FreezePanes = true;
              bool readOnly = false;
              excelApp.DisplayAlerts = false;
              workbook.SaveAs(dlg.FileName, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, readOnly, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing);
              workbook.Close(true, Type.Missing, Type.Missing);
              excelApp.Quit();
              releaseObject(xlR);
              releaseObject(dataSheet);
              releaseObject(sheets);
              releaseObject(workbook);
              releaseObject(excelApp);
              worked = true;
            } catch (Exception e) {
              //a terrible attempt to debug the issue, possible COM Interop investigation
              workbook.Close(false, Type.Missing, Type.Missing);
              excelApp.Quit();
              string m = e.Message;
              MessageBoxResult dialogResult = new MessageBoxResult();
              dialogResult = MessageBox.Show("Investigate/remove " + debug[1] + " " + debug[0] + Environment.NewLine + debug[3] + " ERROR:" + e.Message + Environment.NewLine + "Possible JIRA item polled while it was being changed.", "TRY AGAIN - Fatal error with Excel", MessageBoxButton.OK);

              Dispatcher.Invoke((System.Action)delegate() {
                status.Content = "Failed! Excel got corrupted :(";
              });
              result = false;
            } finally {
              
              releaseObject(xlR);
              releaseObject(dataSheet);
              releaseObject(sheets);
              releaseObject(workbook);
              releaseObject(excelApp);
            }

          }

          disableButtons(false);
          Dispatcher.Invoke((System.Action)delegate() {
            pBar.Visibility = Visibility.Hidden;
            pBar.Value = 0;
            if (worked) {
              status.Content = "Completed!";
            }
          });
          #endregion
        } else {
          disableButtons(false);
          Dispatcher.Invoke((System.Action)delegate() {
            status.Content = "Network failure?! just try again shortly..." + Environment.NewLine + "If this continues, try someplace else.";
            pBar.Value = 0;
            pBar.Visibility = System.Windows.Visibility.Hidden;
          });
        }
      }//if file name chosen
    }

    private void disableButtons(bool disabled) {
      Dispatcher.Invoke((System.Action)delegate() {
        if (disabled) {
          createBtn.IsEnabled = false;
          clrBtn.IsEnabled = false;
          customers.IsEnabled = false;
          dtPick.IsEnabled = false;
          ProjectsStr.IsEnabled = false;
          updateBtn.IsEnabled = false;
          ignoreBox.IsEnabled = false;
        } else {
          createBtn.IsEnabled = true;
          clrBtn.IsEnabled = true;
          customers.IsEnabled = true;
          dtPick.IsEnabled = true;
          ProjectsStr.IsEnabled = true;
          updateBtn.IsEnabled = true;
          ignoreBox.IsEnabled = true;
        }
      });
    }

    private void UPWR(string[] pids) {
      string dt = DateTime.Now.Day.ToString("00") + DateTime.Now.Month.ToString("00") + DateTime.Now.Year.ToString().Substring(2, 2);
      Microsoft.Win32.OpenFileDialog ofg = new Microsoft.Win32.OpenFileDialog();
      ofg.DefaultExt = ".xlsx";
      ofg.Filter = "WR Spreadsheet (.xlsx)|*.xlsx";
      Nullable<bool> result = ofg.ShowDialog();

      List<DevT> DevTrackList = new List<DevT>();
      bool worked = false;
      string[] debug = new string[4];
      if (result == true) {
        Dispatcher.Invoke((System.Action)delegate() {
          status.Content = "Analyzing spreadsheet...";
          pBar.Visibility = Visibility.Visible;
          pBar.Value = 0;
        });

        string textSave = ofg.FileName.Substring(0, ofg.FileName.Length - 4) + "txt";
        string dtPicked = "";//2015-01-27
        DateTime selectedDT = new DateTime();
        bool doAll = false;
        Dispatcher.Invoke((System.Action)delegate() {
          doAll = doAllcb.IsChecked.Value;
          if (!doAll) {
            selectedDT = dtPick.SelectedDate.Value;
            string dtDay = dtPick.SelectedDate.Value.Day.ToString("00");
            string dtMonth = dtPick.SelectedDate.Value.Month.ToString("00");
            string dtYear = dtPick.SelectedDate.Value.Year.ToString("0000");
            dtPicked = dtYear + "-" + dtMonth + "-" + dtDay;
          }

        });
        using (StreamWriter sw = File.CreateText(textSave)) {
          if (doAll) {
            sw.WriteLine("All Dates");
          } else {
            sw.WriteLine(dtPicked);
          }
          foreach (string pid in pids) {
            sw.WriteLine(pid);
          }
        }

        #region analyzeExcel
        Excel.Application xlApp = null;
        Excel.Workbook OldBook = null;
        Excel.Worksheet OldSheet = null;
        Excel.Sheets sheets = null;
        Excel.Range xlR = null;
        Excel.Worksheet dpSheet = null;
        try {
          xlApp = new Excel.Application(); ;
          xlApp.DisplayAlerts = false;
          OldBook = xlApp.Workbooks.Open(ofg.FileName, false, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
          OldSheet = OldBook.Sheets[1];
          sheets = OldBook.Sheets;

          string oldversion = OldSheet.get_Range("G1").Value;
          if (oldversion != "Request" + Environment.NewLine + "Status") {
            xlR = OldSheet.get_Range("G1", "I1").EntireColumn;
            xlR.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
            OldSheet.Cells[1, 7] = "Request" + Environment.NewLine + "Status";
            OldSheet.Cells[1, 8] = "Requested" + Environment.NewLine + "Date";
            OldSheet.Cells[1, 9] = "Planned" + Environment.NewLine + "Release Date";
          }


          int OldMR = OldSheet.UsedRange.Rows.Count;
          object[,] OldData = OldSheet.get_Range("A2:R" + OldMR).Value;
          List<string> OldWRs = getColumn(OldData, 5);
          List<string> OldPID = getColumn(OldData, 3);
          string ignoreItem = "qwertyuiop12345";
          Dispatcher.Invoke((System.Action)delegate() {
            status.Content = "Retrieving data from DevTrack...";
            pBar.Visibility = Visibility.Visible;
            pBar.Value = 0;
            ignoreItem = ignoreBox.Text;
          });
          string[] ignoreItems = System.Text.RegularExpressions.Regex.Split(ignoreItem, ", ");

          List<DevT> OpenDevItems = OpenDevT(pids, ignoreItems);
          List<DevT> LatestOpenDevItems = new List<DevT>();
          List<WR> JIRAitems = new List<WR>();


          bool isError = false;

          foreach (DevT DTitem in OpenDevItems) {
            if (DTitem == null)
              isError = true;
          }
          if (isError) {//devtrack connection issues? skipped a beat
            Dispatcher.Invoke((System.Action)delegate() {
              pBar.Visibility = System.Windows.Visibility.Hidden;
              status.Content = "DevTrack possibly busy/unavailable, trying to reconnect." + Environment.NewLine + "Please wait, retries occur ~3seconds..";
            });
            while (isError) {
              Thread.Sleep(3000);//wait for devtrack to cooldown, possible connection/timeout issues.
              OpenDevItems = OpenDevT(pids, ignoreItems);
              for (int i = 0; i < OpenDevItems.Count; i++) {
                if (OpenDevItems[i] == null) {
                  isError = true;
                  break;
                } else {
                  isError = false;
                }
              }
            }
          }



          if (!doAll) {
            Parallel.ForEach(OpenDevItems, DTitem => {
              if (DTitem.createdDate >= selectedDT) {
                LatestOpenDevItems.Add(DTitem);
              }
            });
          } else {
            LatestOpenDevItems = OpenDevItems;
          }
          List<string> nowClosed = getNowClosed(LatestOpenDevItems, OldWRs);

          List<string> urls = makeUrl(LatestOpenDevItems);
          List<string> closedURLs = mkoldURLs(nowClosed);
          if (nowClosed.Count > 0) {
            urls.AddRange(closedURLs);
          }



          double j = 100.0;
          double k = j / urls.Count();
          Dispatcher.Invoke((System.Action)delegate() {
            status.Content = "Querying a lot of JIRA...";
            pBar.Visibility = Visibility.Visible;
            pBar.Value = 0;
          });
          
          //int speed = 987;//get wait

          //foreach (string qry in urls) {
          //  JIRAitems.Add(pJSON(qry));
          //  Dispatcher.Invoke((System.Action)delegate() {
          //    pBar.Value += k;
          //  });
          //  Thread.Sleep(speed);
          //}

          Parallel.ForEach(urls, qry => {
            JIRAitems.Add(pJSON(qry));
            Dispatcher.Invoke((System.Action)delegate() {
              pBar.Value += k;
            });
          });

          Dispatcher.Invoke((System.Action)delegate() {
            status.Content = "Parsing JIRA data...";
          });
          List<WRQuery> JItemsC = new List<WRQuery>();
          foreach (WR dd in JIRAitems) {
            foreach (Jissues info in dd.issues) {
                JItemsC.Add(new WRQuery(info, pids));
             
              
            }
          }
          Dispatcher.Invoke((System.Action)delegate() {
            status.Content = "Finding JIRA errors...";
          });
          List<string> notInJIRA = new List<string>();
          foreach (DevT lodi in LatestOpenDevItems) {
            if (DevWRExists(JItemsC, lodi.WR) < 0) {
              notInJIRA.Add(lodi.WR);
            }
          }

          Dispatcher.Invoke((System.Action)delegate() {
            status.Content = "Wrestling Excel...";
            pBar.Value = 0;
          });


          if (JIRAitems.Count != 0) {
            List<string> cautions = new List<string>();
            string dtNow = DateTime.Now.Month.ToString("00") + "-" + DateTime.Now.Day.ToString("00") + "-" + DateTime.Now.Year.ToString().Substring(2, 2);
            if (!doAll) {
              OldSheet.Name = dtPicked + " and newer";
            } else {
              OldSheet.Name = "All";
            }

            int nXLcount = LatestOpenDevItems.Count();
            double iB = 100.0;
            double iJ = iB / nXLcount;
            int rowCount = 2; //start at row 2
            int insOffset = 0;
            int maxR = OldData.GetLength(0);
            foreach (DevT DT in LatestOpenDevItems) { //foreach found valid devtrack item
              debug[0] = DT.WR;
              debug[1] = DT.Project;
              int xWR = DevWRExists(JItemsC, DT.WR); //find it in the JIRA data
              string wSuite = "WR NOT IN JIRA!";
              string wSummary = "WR NOT IN JIRA!";
              string wCurRel = "WR NOT IN JIRA!";
              string wIT = "WR NOT IN JIRA!";
              string wProd = "WR NOT IN JIRA!";
              string wStatus = "WR NOT IN JIRA!";
              string rStatus = "WR NOT IN JIRA!";
              string rDate = "WR NOT IN JIRA!";
              string dDate = "WR NOT IN JIRA!";
              string wDesc = "WR NOT IN JIRA! Contact Development for assistance with this WR.";
              if (xWR >= 0) { //if it exists in the JIRA data, update the default strings
                wSuite = JItemsC[xWR].Suite;
                wSummary = JItemsC[xWR].Summary;
                rStatus = JItemsC[xWR].RequestedStatus;
                rDate = JItemsC[xWR].RequestedDate;
                dDate = JItemsC[xWR].DueDate;
                wCurRel = JItemsC[xWR].CurrentRelease;
                wIT = JItemsC[xWR].IssueType;
                wStatus = JItemsC[xWR].Resolution + Environment.NewLine + JItemsC[xWR].Status;
                wProd = JItemsC[xWR].Product;
                wDesc = JItemsC[xWR].Description;
              }
              int jets = getSIndex(OldWRs, OldPID, DT.WR, DT.Project);
              debug[3] = jets.ToString();
              if (jets < 0) {
                //new item
                string cdate = DT.createdDate.Month + "-" + DT.createdDate.Day + "-" + DT.createdDate.Year;
                string[] newLine = { dtNow, "No", DT.Project, wSuite, DT.WR, wStatus, rStatus, rDate, dDate, cdate, wSummary, "'" + wDesc, wProd, wCurRel, DT.IV, DT.IVseverity, DT.NeedByEvent, DT.NeedByDate };
                object[,] xlNewLine = new object[1, newLine.Length];
                for (int i = 0; i < newLine.Length; i++) {
                  xlNewLine[0, i] = newLine[i];
                }
                //insert row at end
                xlR = OldSheet.get_Range("A" + (maxR + 2), "A" + (maxR + 2)).EntireRow;
                xlR.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                xlR = OldSheet.Range["A" + (maxR + 2), getAlpha(newLine.Length) + (maxR + 2)];
                xlR.Interior.ColorIndex = 0;
                xlR.Value2 = xlNewLine;
                maxR++;
              } else {
                //existing item
                //check if updates occured
                //-update item, highlight yellow
                //

                string[] oldRow = getRow(OldData, jets + 1).ToArray();
                bool updated = false;
                int[] colData = new int[] { 4, 6, 7, 8, 9, 11, 13, 14, 16, 17, 18 };
                string[] colDx = new string[] { wSuite, wStatus, rStatus, rDate, dDate, wSummary, wProd, wCurRel, DT.IVseverity, DT.NeedByEvent, DT.NeedByDate };
                for (int i = 0; i < colData.Length; i++) {
                  string undated = undate(oldRow[colData[i] - 1]);
                  if (undate(colDx[i]) != undated) {
                    if (undated != "") {
                      string ttemp = "Was: " + undated + Environment.NewLine + Environment.NewLine + "Now: " + colDx[i];
                      OldSheet.Range[getAlpha(colData[i]) + (insOffset + jets + 2), Type.Missing].Interior.ColorIndex = 6;
                      OldSheet.Cells[insOffset + jets + 2, colData[i]] = ttemp;
                      updated = true;
                    } else {
                      string ttemp = colDx[i];
                      OldSheet.Cells[insOffset + jets + 2, colData[i]] = ttemp;
                    }
                  }
                }
                if (updated) {
                  OldSheet.Cells[insOffset + jets + 2, 2] = "Yes";
                }
              }

              //progress bar
              rowCount++;
              Dispatcher.Invoke((System.Action)delegate() {
                pBar.Value += iJ;
              });
            }
            foreach (string closedWR in nowClosed) {
              int xWR = DevWRExists(JItemsC, closedWR);
              string rStatus = "WR NOT IN JIRA!";
              string rDate = "WR NOT IN JIRA!";
              string dDate = "WR NOT IN JIRA!";
              string wSuite = "WR NOT IN JIRA!";
              string wSummary = "WR NOT IN JIRA!";
              string wCurRel = "WR NOT IN JIRA!";
              string wIT = "WR NOT IN JIRA!";
              string wProd = "WR NOT IN JIRA!";
              string wStatus = "WR NOT IN JIRA!";
              string wDesc = "WR NOT IN JIRA! Contact Development for assistance with this WR.";
              if (xWR >= 0) { //if it exists in the JIRA data, update the default strings
                wSuite = JItemsC[xWR].Suite;
                wSummary = JItemsC[xWR].Summary;
                wCurRel = JItemsC[xWR].CurrentRelease;
                rStatus = JItemsC[xWR].RequestedStatus;
                rDate = JItemsC[xWR].RequestedDate;
                dDate = JItemsC[xWR].DueDate;
                wIT = JItemsC[xWR].IssueType;
                wStatus = JItemsC[xWR].Resolution + Environment.NewLine + JItemsC[xWR].Status;
                wProd = JItemsC[xWR].Product;
                wDesc = JItemsC[xWR].Description;
              }
              int jets = getIndex(OldWRs, closedWR);
              debug[3] = jets.ToString();
              string[] oldRow = getRow(OldData, jets + 1).ToArray();
              int[] colData = new int[] { 4, 6, 7, 8, 9, 11, 13, 14 };
              string[] colDx = new string[] { wSuite, wStatus, rStatus, rDate, dDate, wSummary, wProd, wCurRel };
              for (int i = 0; i < colData.Length; i++) {
                string undated = undate(oldRow[colData[i] - 1]);
                if (undate(colDx[i]) != undated) {
                  if (undated != "") {
                    string ttemp = "Was: " + undated + Environment.NewLine + Environment.NewLine + "Now: " + colDx[i];
                    OldSheet.Range[getAlpha(colData[i]) + (insOffset + jets + 2), Type.Missing].Interior.ColorIndex = 6;
                    OldSheet.Cells[insOffset + jets + 2, colData[i]] = ttemp;

                  } else {
                    string ttemp = colDx[i];
                    OldSheet.Cells[insOffset + jets + 2, colData[i]] = ttemp;
                  }
                }
              }
              string xtemp = oldRow[9].Split(' ')[0];
              OldSheet.Cells[insOffset + jets + 2, 10] = "Was: " + xtemp + Environment.NewLine + Environment.NewLine + "But it was not found" + Environment.NewLine + "Check dates, closure," + Environment.NewLine + "or linked status.";
              OldSheet.Range["J" + (insOffset + jets + 2), Type.Missing].Interior.ColorIndex = 6;
              OldSheet.Cells[insOffset + jets + 2, 2] = "Yes";
            }



          }
        #endregion



          int maxRf = OldSheet.UsedRange.Rows.Count;
          object[,] arr = OldSheet.get_Range("C2:E" + maxRf).Value;
          List<string> ProjectRow = new List<string>();
          List<string> ProjectWR = new List<string>();
          for (int i = 1; i < maxRf; i++) {
            ProjectRow.Add(arr[i, 1].ToString());
            ProjectWR.Add(arr[i, 3].ToString());
          }
          List<string> duplicatedWRs = ProjectWR.GroupBy(x => x)
            .Where(group => group.Count() > 1)
            .Select(group => group.Key).ToList();

          if (duplicatedWRs.Count > 0) {
            if (sheets.Count > 1) {
              xlApp.DisplayAlerts = false;
              OldBook.Sheets[2].Delete();
            }
            dpSheet = (Worksheet)sheets.Add(Type.Missing, sheets[1], Type.Missing, Type.Missing);
            dpSheet.Name = "Updated Duplicate WRs";
            int rowToPrint = 2;
            dpSheet.Cells[1, 1] = "Duplicated" + Environment.NewLine + "WRs Found";
            dpSheet.Cells[1, 2] = "Projects found linked";
            dpSheet.Cells[1, 3] = "These WRs did NOT get updated correctly, and will require a manual review!";
            foreach (string dWRx in duplicatedWRs) {
              debug[3] = rowToPrint.ToString();
              string ProjectsFound = "";
              for (int i = 0; i < ProjectWR.Count; i++) {
                if (dWRx == ProjectWR[i]) {
                  if (ProjectsFound == "") {
                    ProjectsFound += ProjectRow[i];
                  } else {
                    ProjectsFound += Environment.NewLine + ProjectRow[i];
                  }
                }
              }
              //print this WR
              dpSheet.Cells[rowToPrint, 1] = dWRx;
              dpSheet.Cells[rowToPrint, 2] = ProjectsFound;
              rowToPrint++;
            }
            dpSheet.get_Range("A:A", Type.Missing).EntireColumn.ColumnWidth = 16;
            dpSheet.get_Range("B:B", Type.Missing).EntireColumn.ColumnWidth = 32;
            dpSheet.get_Range("A:B", Type.Missing).EntireColumn.VerticalAlignment = XlVAlign.xlVAlignCenter;
            dpSheet.get_Range("A:B", Type.Missing).EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            dpSheet.get_Range("A1", "B1").Cells.Font.Bold = true;
            dpSheet.Select();
            //
          } else {
            if (sheets.Count > 1) {
              xlApp.DisplayAlerts = false;
              OldBook.Sheets[2].Delete();
            }
          }

          //Data Validation
          string validS = "Urgent,High,Med,Low,SCT";//MATCH WRPR above
          string validW = "Urgent,High,Med,Low";
          //string validC = "High,Med,Low,Not required";
          string validU = "Yes,No";
          OldSheet.get_Range("B2", "Z" + (maxRf - 1)).Validation.Delete(); //gotta clear it before we can set it.
          OldSheet.get_Range("B2", "B" + (maxRf - 1)).Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, validU, Type.Missing);
          OldSheet.get_Range("S2", "S" + (maxRf - 1)).Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, validS, Type.Missing);
          OldSheet.get_Range("W2", "W" + (maxRf - 1)).Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, validW, Type.Missing);
          OldSheet.get_Range("U2", "U" + (maxRf - 1)).Validation.Add(XlDVType.xlValidateList, XlDVAlertStyle.xlValidAlertStop, XlFormatConditionOperator.xlBetween, validW, Type.Missing);


          worked = true;
          OldBook.SaveAs(ofg.FileName, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing);
          OldBook.Close(true, Type.Missing, Type.Missing);
          xlApp.Quit();
        } catch (Exception e) {
          OldBook.Close(false, Type.Missing, Type.Missing);
          xlApp.Quit();
          if(e.Message=="Object reference not set to an instance of an object."){
            MessageBoxResult dialogResult = new MessageBoxResult();
            dialogResult = MessageBox.Show("Formatting issue " + debug[1] + " " + debug[0] + Environment.NewLine + debug[3] + " ERROR:" + e.Message + Environment.NewLine + "Likely blank rows at bottom of table." +Environment.NewLine + "Try deleting several entire rows below the table.", "TRY AGAIN - Fatal error with Excel", MessageBoxButton.OK);
            Dispatcher.Invoke((System.Action)delegate() {
              status.Content = "Failed! Excel is garbage :(";
            });
          } else {
            MessageBoxResult dialogResult = new MessageBoxResult();
            dialogResult = MessageBox.Show("Investigate/remove " + debug[1] + " " + debug[0] + Environment.NewLine + debug[3] + " ERROR:" + e.Message + Environment.NewLine + "Possible JIRA interference!", "TRY AGAIN - Fatal error with Excel", MessageBoxButton.OK);
            Dispatcher.Invoke((System.Action)delegate() {
              status.Content = "Failed! Excel is garbage :(";
            });
          }
          
          result = false;
        } finally {         
          releaseObject(xlR);
          releaseObject(dpSheet);
          releaseObject(sheets);
          releaseObject(OldSheet);
          releaseObject(OldBook);
          releaseObject(xlApp);
        }

      }
      Dispatcher.Invoke((System.Action)delegate() {
        pBar.Visibility = Visibility.Hidden;
        pBar.Value = 0;

        if (worked) {
          status.Content = "Completed! Please close the application.";
        }
      });
    }//if file name chosen

    private void verifyAndUpdate(object sender, DoWorkEventArgs e) {
      List<string> customerList = (List<string>)e.Argument;
      List<string> updatedList = new List<string>();
      string projectQry = "";
      disableButtons(true);
      Dispatcher.Invoke((System.Action)delegate() {
        projectQry = ProjectsStr.Text;
        updatedList = customers.SelectedItems.OfType<string>().ToList();
      });
      if (projectQry != "") {
        string[] pQOs = projectQry.Split(',');
        foreach (string projectQ in pQOs) {
          wildcard chaos = new wildcard(projectQ, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
          foreach (string projectCode in customerList) {
            if (chaos.IsMatch(projectCode))
              if (getIndex(updatedList, projectCode) < 0) {
                updatedList.Add(projectCode);
              }
          }
        }
      } else {
        updatedList = customerList;
      }

      string[] pid = updatedList.ToArray();
      MessageBoxResult dialogResult = new MessageBoxResult();
      if (pid.Length > 0 && pid.Length < 11) {
        dialogResult = MessageBox.Show("This will update(overwrite) a report with: " + Environment.NewLine + string.Join(" & " + Environment.NewLine, pid), "WR Priority Report", MessageBoxButton.OKCancel);
      } else if (pid.Length >= 11) {
        dialogResult = MessageBox.Show("This will update(overwrite) a report for " + pid.Length + " projects", "WR Priority Report", MessageBoxButton.OKCancel);
      } else {
        dialogResult = MessageBox.Show("No projects found to match", "WR Priority Report", MessageBoxButton.OK);
      }
      if (dialogResult == MessageBoxResult.OK && pid.Length > 0) {
        //WRPR(pid);
        //
        UPWR(pid);
      } else {
        disableButtons(false);
      }
    }

    private void updateBtn_Click(object sender, RoutedEventArgs e) {
      //check projects
      List<string> items = new List<string>();
      if (ProjectsStr.Text == "") {
        foreach (var item in customers.SelectedItems) {
          items.Add(item.ToString());
        }
      } else {
        items = customers.Items.OfType<string>().ToList();
      }
      if (items.Count != 0) {
        BackgroundWorker bw1 = new BackgroundWorker();
        bw1.DoWork += new DoWorkEventHandler(verifyAndUpdate);
        bw1.RunWorkerAsync(items);
      } else {
        MessageBoxResult dialogResult = MessageBox.Show("Please select a project", "WR Priority Report", MessageBoxButton.OK);
      }

    }

    private void doAllcb_Checked(object sender, RoutedEventArgs e) {

    }

  }
}
