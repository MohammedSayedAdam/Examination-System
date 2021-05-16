using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Diagnostics;
using MicrosoftEdgecls;
using System.Runtime.InteropServices;
using System.Windows.Automation;

namespace Eximination_System
{
    public partial class Form1 : Form
    {

        //create new json file nameed bookmark to jsutify this project
        // when i reset i wil remove the file named bookmarks if exsist and write this json file
        // trace why the code delete the file when i  call delete from favorite function
        //https://www.newtonsoft.com/json/help/html/SerializingJSONFragments.htm
        //
        //string Jbookmark = @"{
        //                   'checksum': 'b7303559d2f17060bd1eb78b6aab2140',
        //                   'roots': {
        //                      'bookmark_bar': {
        //                         'children': [  ],
        //                         'date_added': '13263223872257087',
        //                         'date_modified': '13263223876892461',
        //                         'guid': '00000000-0000-4000-a000-000000000002',
        //                         'id': '1',
        //                         'name': 'Favorites bar',
        //                         'source': 'unknown',
        //                         'type': 'folder'
        //                      },
        //                      'other': {
        //                         'children': [  ],
        //                         'date_added': '13263223872257093',
        //                         'date_modified': '0',
        //                         'guid': '00000000-0000-4000-a000-000000000003',
        //                         'id': '2',
        //                         'name': 'Other favorites',
        //                         'source': 'unknown',
        //                         'type': 'folder'
        //                      },
        //                      'synced': {
        //                         'children': [  ],
        //                         'date_added': '13263223872257095',
        //                         'date_modified': '0',
        //                         'guid': '00000000-0000-4000-a000-000000000004',
        //                         'id': '3',
        //                         'name': 'Mobile favorites',
        //                         'source': 'unknown',
        //                         'type': 'folder'
        //                      }
        //                   },
        //                   'version': 1
        //                }
        //               ";
        public void getquickaccess()
        {
            var path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var subFolderPath = Path.Combine(path, "MyAssignment1");
            Process.Start("Quick access");
        }
        [DllImport("user32")]
        public static extern IntPtr GetDesktopWindow();
        public bool FindMsEdgeWindow(string TabName)
        {
            Process[] procsEdge = Process.GetProcessesByName("msedge");
            if (procsEdge.Length <= 0)
            {
                return false;
            }
            else
            {
                int Wcounter = 0;
                bool ok = false;
                foreach (Process proc in procsEdge)
                {
                    //the Edge process must have a window
                    if (proc.MainWindowHandle != IntPtr.Zero)
                    {
                       
                        AutomationElement root = AutomationElement.FromHandle(proc.MainWindowHandle);
                        TreeWalker treewalker = TreeWalker.ControlViewWalker;
                        AutomationElement rootParent = treewalker.GetParent(root);
                        Condition condWindow = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
                        Condition condNewTab = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);
                        AutomationElementCollection edges = rootParent.FindAll(TreeScope.Children, condWindow);
                        foreach (AutomationElement e in edges)
                        {
                            //check if the root element is named with *Edge*
                            //YouTube and 2 more pages - Profile 1 - Microsoft​ Edg
                            //Google - Profile 1 - Microsoft​ Edge
                            if (e.Current.Name.Contains("Microsoft​ Edge"))
                            {

                                Wcounter++;
                               
                                foreach (AutomationElement tabitem in e.FindAll(TreeScope.Descendants, condNewTab))
                                {
                                    if (tabitem.Current.Name.Contains(TabName))
                                    {
                                        ok =  true;
                                    }
                                }
                                if (e.Current.Name.Contains(TabName))
                                {
                                    ok = true;
                                }
                            }
                        }

                    }


                }
                if(Wcounter >1 && ok)
                {
                    return true;
                }
            }
            return false;
        }
        public bool MsEdgeInPrivate(string TabName)
        {
            Process[] procsEdge = Process.GetProcessesByName("msedge");
            if (procsEdge.Length <= 0)
            {
                return false;
            }
            else
            {
                foreach (Process proc in procsEdge)
                {
                    //the Edge process must have a window
                    if (proc.MainWindowHandle != IntPtr.Zero)
                    {
                        AutomationElement root = AutomationElement.FromHandle(proc.MainWindowHandle);
                        TreeWalker treewalker = TreeWalker.ControlViewWalker;
                        AutomationElement rootParent = treewalker.GetParent(root);
                        Condition condWindow = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
                        Condition condNewTab = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);
                        AutomationElementCollection edges = rootParent.FindAll(TreeScope.Children, condWindow);
                        foreach (AutomationElement e in edges)
                        {
                            //In - Private browsing - Bing - [InPrivate] - Microsoft​ Edge
                            if (e.Current.Name.Contains("Microsoft​ Edge") && e.Current.Name.Contains("InPrivate") && e.Current.Name.Contains(TabName)) 
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }
        public bool FindMsEdgeTabs(string TabName)
        {
            ////if(FindMsEdgeTabs("New tab"))
            ////{
            ////    MessageBox.Show("he is sucesseded");
            ////}
            //if (FindMsEdgeTabs("YouTube"))
            //{
            //    MessageBox.Show("he is sucesseded");
            //}
            //else
            //{
            //    MessageBox.Show("he is failed");
            //}
            //HideTaskBar();
            // CheckNewWindowMSEdge();
            ////// Process.GetCurrentProcess();
            //https://stackoverflow.com/questions/40070703/how-to-get-a-list-of-open-tabs-from-chrome-c-sharp


            Process[] procsEdge = Process.GetProcessesByName("msedge");
            if (procsEdge.Length <= 0)
            {
                return false;
            }
            else
            {
                foreach (Process proc in procsEdge)
                {
                    //the Edge process must have a window
                    if (proc.MainWindowHandle != IntPtr.Zero)
                    {
                        AutomationElement root = AutomationElement.FromHandle(proc.MainWindowHandle);
                        TreeWalker treewalker = TreeWalker.ControlViewWalker;
                        AutomationElement rootParent = treewalker.GetParent(root);
                        Condition condWindow = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window);
                        Condition condNewTab = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem);
                        AutomationElementCollection edges = rootParent.FindAll(TreeScope.Children, condWindow);
                        foreach (AutomationElement e in edges)
                        {
                            //check if the root element is named with *Edge*
                            //YouTube and 2 more pages - Profile 1 - Microsoft​ Edg
                            //Google - Profile 1 - Microsoft​ Edge
                            //In - Private browsing - Bing - [InPrivate] - Microsoft​ Edge
                            if (e.Current.Name.Contains("Microsoft​ Edge")) // to ensure that he open it in tab and not window
                            {
                                if (e.Current.Name.Contains("more"))
                                {
                                    foreach (AutomationElement tabitem in e.FindAll(TreeScope.Descendants, condNewTab))
                                    {
                                        if (tabitem.Current.Name.Contains(TabName))
                                        {
                                            return true;
                                        }
                                    }
                                } 
                                if (e.Current.Name.Contains(TabName)&& e.Current.Name.Contains("more"))
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }
            }
            return false;
        }
        [DllImport("user32.dll")]
        private static extern int FindWindow(string className, string windowText);
        [DllImport("user32.dll")]
        private static extern int ShowWindow(int hwnd, int command);
        private const int SW_HIDE = 0;
        private const int SW_SHOW = 1;
        private bool ShowTaskBar()
        {


            int hwnd = FindWindow("Shell_TrayWnd", "");
            ShowWindow(hwnd, 0);
            return true;
        }
        private bool HideTaskBar()
        {


            int hwnd = FindWindow("Shell_TrayWnd", "");
            ShowWindow(hwnd, 0);
            return true;
        }
        public MsEdgeOperations edge = new MicrosoftEdgecls.MsEdgeOperations();
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
        public void CheckNewWindowMSEdge()
        {
            Process[] processlist = Process.GetProcesses();

            foreach (Process process in processlist)
            {
                if (!String.IsNullOrEmpty(process.MainWindowTitle))
                {
                    //if(process.ProcessName == "msedge")
                   MessageBox.Show( process.ProcessName+" "+ process.MainWindowTitle);
                }
            }
        }
        public void RemoveFavoriteFromMicroedge()
        {
            string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string[] PCname;
            JObject o2;
            string newJson = "";
            PCname = userName.Split('\\');
            string path = @"C:\Users\" + PCname[1] + @"\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks";
            using (StreamReader file = File.OpenText(path))
            using (JsonTextReader reader = new JsonTextReader(file))
            {
                o2 = (JObject)JToken.ReadFrom(reader);

                foreach (JObject item in o2["roots"]["bookmark_bar"]["children"].Children<JObject>())
                {
                    if (item["name"] != null)
                        item["url"].Parent.Remove();
                    newJson = JsonConvert.SerializeObject(o2);
                }
                foreach (JObject item in o2["roots"]["other"]["children"].Children<JObject>())
                {
                    if (item["name"] != null)
                        item["url"].Parent.Remove();
                    newJson = JsonConvert.SerializeObject(o2);
                }
            }
            File.WriteAllText(path, newJson);
        }
        private void btn_submit_Click(object sender, EventArgs e)
        {
            if (MsEdgeInPrivate("skysports"))
            {
                MessageBox.Show("he is sucesseded");
            }
            else
            {
                MessageBox.Show("he is failed");
            }
            //try
            //{
            //    edge.Invoke(textBox1.Text);

            //}
            //catch (Exception ex)
            //{

            //    MessageBox.Show(ex.Data.ToString());
            //}
            //getquickaccess();
        }
        //private void url()
        //{
        //    UIAutomationClient.IUIAutomationElement rootElement = uiAutomation.GetRootElement();

        //    int propertyName = 30005; // UIA_NamePropertyId
        //    int propertyAutomationId = 30011; // UIA_AutomationIdPropertyId
        //    int propertyClassName = 30012; // UIA_ClassNamePropertyId
        //    int propertyNativeWindowHandle = 30020; // UIA_NativeWindowHandlePropertyId

        //    // Get the main Edge element, which is a direct child of the UIA root element.
        //    // For this test, assume that the Edge element is the only element with an
        //    // AutomationId of "TitleBar".
        //    string edgeAutomationId = "TitleBar";

        //    UIAutomationClient.IUIAutomationCondition condition =
        //        uiAutomation.CreatePropertyCondition(
        //            propertyAutomationId, edgeAutomationId);

        //    // Have the window handle cached when we find the main Edge element.
        //    UIAutomationClient.IUIAutomationCacheRequest cacheRequestNativeWindowHandle = uiAutomation.CreateCacheRequest();
        //    cacheRequestNativeWindowHandle.AddProperty(propertyNativeWindowHandle);

        //    UIAutomationClient.IUIAutomationElement edgeElement =
        //        rootElement.FindFirstBuildCache(
        //            UIAutomationClient.TreeScope.TreeScope_Children,
        //            condition,
        //            cacheRequestNativeWindowHandle);

        //    if (edgeElement != null)
        //    {
        //        IntPtr edgeWindowHandle = edgeElement.CachedNativeWindowHandle;

        //        // Next find the element whose name is the url of the loaded page. And have
        //        // the name of the element related to the url cached when we find the element.
        //        UIAutomationClient.IUIAutomationCacheRequest cacheRequest =
        //            uiAutomation.CreateCacheRequest();
        //        cacheRequest.AddProperty(propertyName);

        //        // For this test, assume that the element with the url is the first descendant element
        //        // with a ClassName of "Internet Explorer_Server".
        //        string urlElementClassName = "Internet Explorer_Server";

        //        UIAutomationClient.IUIAutomationCondition conditionUrl =
        //            uiAutomation.CreatePropertyCondition(
        //                propertyClassName,
        //                urlElementClassName);

        //        UIAutomationClient.IUIAutomationElement urlElement =
        //            edgeElement.FindFirstBuildCache(
        //                UIAutomationClient.TreeScope.TreeScope_Descendants,
        //                conditionUrl,
        //                cacheRequest);

        //        string url = urlElement.CachedName;

        //        // Next find the title of the loaded page. First find the list of 
        //        // tabs shown at the top of Edge.
        //        string tabsListAutomationId = "TabsList";

        //        UIAutomationClient.IUIAutomationCondition conditionTabsList =
        //            uiAutomation.CreatePropertyCondition(
        //                propertyAutomationId, tabsListAutomationId);

        //        UIAutomationClient.IUIAutomationElement tabsListElement =
        //            edgeElement.FindFirst(
        //                UIAutomationClient.TreeScope.TreeScope_Descendants,
        //                conditionTabsList);

        //        // Find which of those tabs is selected. (It should be possible to 
        //        // cache the Selection pattern with the above call, and that would
        //        // avoid one cross-process call here.)
        //        int selectionPatternId = 10001; // UIA_SelectionPatternId
        //        IUIAutomationSelectionPattern selectionPattern =
        //            tabsListElement.GetCurrentPattern(selectionPatternId);

        //        // For this test, assume there's always one selected item in the list.
        //        UIAutomationClient.IUIAutomationElementArray elementArray = selectionPattern.GetCurrentSelection();
        //        string title = elementArray.GetElement(0).CurrentName;

        //        // Now show the title, url and window handle.
        //        MessageBox.Show(
        //            "Page title: " + title +
        //            "\r\nURL: " + url +
        //            "\r\nhwnd: " + edgeWindowHandle);
        //    }
        //}
        private void button1_Click(object sender, EventArgs e)
        {
            
            try
            {
                if (edge.GetResult(textBox1.Text))
                {
                    MessageBox.Show("he is sucesseded");
                }
                else
                    MessageBox.Show("he is Failed");
                edge.ResetToDefault(textBox1.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

    }

}



//System.Diagnostics.Process.Start("microsoft-edge:"); E:\ADAM\INSPIREPQ\INSPIREPQ\bin\Debug\INSPIREPQ_V8\    E:\ADAM\INSPIREPQ\bin\Debug
//System.Diagnostics.Process.Start("microsoft-edge:http://www.google.com");
//string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
//string[] PCname;
//JObject o2;
//PCname = userName.Split('\\');
//string path = @"C:\Users\" + PCname[1] + @"\AppData\Local\Microsoft\Edge\User Data\Default\Bookmarks";


//// to fill data in an obj in class jsonobj
//JsonObj obj = new JsonObj();
//obj.date_added = DateTime.Now;
//obj.type = "url";
//obj.name = "Coding";
//obj.url = @"https://www.urionlinejudge.com.br/";
//obj.show_icon = false;

//string newJson;
//using (StreamReader file = File.OpenText(path))
//using (JsonTextReader reader = new JsonTextReader(file))
//{
//    o2 = (JObject)JToken.ReadFrom(reader);
//    string data = o2["roots"]["bookmark_bar"]["children"].ToString();

//    foreach (var item in o2["roots"]["bookmark_bar"]["children"])
//    {
//        MessageBox.Show(item["name"].ToString());
//    }
//    JArray ite = (JArray)o2["roots"]["bookmark_bar"]["children"];
//    ite.Add(JToken.FromObject(obj));
//    newJson = JsonConvert.SerializeObject(o2);
//}
//File.WriteAllText(path, newJson);

//try
//{
//    // all.Invoke(textBox1.Text);
//    if (all.Invoke(textBox1.Text))
//    {

//    }
//    else
//    {

//        MessageBox.Show("NO");
//    }
//    // all.ResetToDefault(textBox1.Text);
//}
//catch (Exception)
//{
//    MessageBox.Show("You Must Enter QuestionID");
//}

