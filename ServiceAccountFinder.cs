// Copyright (C) 2025 Akil Woolfolk Sr. 
// All Rights Reserved
// All the changes released under the MIT license as the original code.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Management;
using System.Globalization;
using Microsoft.Win32;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.DirectoryServices;

namespace ServiceAccountFinder
{
    class SAFMain
    {
        struct CMDArguments
        {
            public bool bParseCmdArguments;
        }

        static ManagementObjectCollection funcSysQueryData(string sysQueryString, string sysServerName)
        {

            // [Comment] Connect to the server via WMI
            System.Management.ConnectionOptions objConnOptions = new System.Management.ConnectionOptions();
            string strServerNameforWMI = "\\\\" + sysServerName + "\\root\\cimv2";

            // [DebugLine] Console.WriteLine("Construct WMI scope...");
            System.Management.ManagementScope objManagementScope = new System.Management.ManagementScope(strServerNameforWMI, objConnOptions);

            // [DebugLine] Console.WriteLine("Construct WMI query...");
            System.Management.ObjectQuery objQuery = new System.Management.ObjectQuery(sysQueryString);
            //if (objQuery != null)
            //    Console.WriteLine("objQuery was created successfully");

            // [DebugLine] Console.WriteLine("Construct WMI object searcher...");
            System.Management.ManagementObjectSearcher objSearcher = new System.Management.ManagementObjectSearcher(objManagementScope, objQuery);
            //if (objSearcher != null)
            //    Console.WriteLine("objSearcher was created successfully");

            // [DebugLine] Console.WriteLine("Get WMI data...");

            System.Management.ManagementObjectCollection objReturnCollection = null;

            try
            {
                objReturnCollection = objSearcher.Get();
                return objReturnCollection;
            }
            catch (SystemException ex)
            {
                // [DebugLine] System.Console.WriteLine("{0} exception caught here.", ex.GetType().ToString());
                string strRPCUnavailable = "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)";
                // [DebugLine] System.Console.WriteLine(ex.Message);
                if (ex.Message == strRPCUnavailable)
                {
                    Console.WriteLine("WMI: Server unavailable");
                }

                // Next line will return an object that is equal to null
                return objReturnCollection;
            }
        }

        static bool funcPingServer(string strServerName)
        {
            bool bPingSuccess = false;
            //bool bWMISuccess = false;

            // [DebugLine] Console.WriteLine("Contact start for {0}: {1}", strServerName, DateTime.Now.ToLocalTime().ToString("MMddyyy HH:mm:ss"));

            //string strServerNameforWMI = "";

            // [Comment] Ping the server
            // [DebugLine] Console.WriteLine(); // Helper line just to make output clearer
            // [DebugLine] Console.WriteLine("Ping attempt for: " + strServerName);

            try
            {
                System.Net.NetworkInformation.Ping objPing1 = new System.Net.NetworkInformation.Ping();
                System.Net.NetworkInformation.PingReply objPingReply1 = objPing1.Send(strServerName);
                if (objPingReply1.Status.ToString() != "TimedOut")
                {
                    // [DebugLine] Console.WriteLine("Ping Reply: " + objPingReply1.Address + "     RTT: " + objPingReply1.RoundtripTime);
                    bPingSuccess = true;
                }
                else
                {
                    // [DebugLine] Console.WriteLine("Ping Reply: " + objPingReply1.Status);
                    bPingSuccess = false;
                }
            }
            catch (SystemException ex)
            {
                // [DebugLine] System.Console.WriteLine("{0} exception caught here.", ex.GetType().ToString());
                string strPingError = "An exception occurred during a Ping request.";
                // [DebugLine] System.Console.WriteLine(ex.Message);
                if (ex.Message == strPingError)
                {
                    Console.WriteLine("Ping Error. No ip address was found during name resolution.");
                }

                bPingSuccess = false;
            }

            //// [Comment] Connect to the server via WMI
            //// [DebugLine] Console.WriteLine(); // Helper line just to make output clearer
            //Console.WriteLine("WMI connection attempt for: " + strServerName);

            //System.Management.ConnectionOptions objConnOptions = new System.Management.ConnectionOptions();
            //strServerNameforWMI = "\\\\" + strServerName + "\\root\\cimv2";

            //// [DebugLine] Console.WriteLine("Construct WMI scope...");
            //System.Management.ManagementScope objManagementScope = new System.Management.ManagementScope(strServerNameforWMI, objConnOptions);
            //// [DebugLine] Console.WriteLine("Construct WMI query...");
            //System.Management.ObjectQuery objQuery = new System.Management.ObjectQuery("select * from Win32_ComputerSystem");
            //// [DebugLine] Console.WriteLine("Construct WMI object searcher...");
            //System.Management.ManagementObjectSearcher objSearcher = new System.Management.ManagementObjectSearcher(objManagementScope, objQuery);
            //Console.WriteLine("Get WMI data...");

            //try
            //{
            //    System.Management.ManagementObjectCollection objObjCollection = objSearcher.Get();

            //    foreach (System.Management.ManagementObject objMgmtObject in objObjCollection)
            //    {
            //        Console.WriteLine("Hostname: " + objMgmtObject["Caption"].ToString());
            //        bWMISuccess = true;
            //    }
            //}
            //catch (SystemException ex)
            //{
            //    // [DebugLine] System.Console.WriteLine("{0} exception caught here.", ex.GetType().ToString());
            //    string strRPCUnavailable = "The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)";
            //    // [DebugLine] System.Console.WriteLine(ex.Message);
            //    if (ex.Message == strRPCUnavailable)
            //    {
            //        Console.WriteLine("WMI: Server unavailable");
            //    }
            //    bWMISuccess = false;
            //}

            // [DebugLine] Console.WriteLine("Contact stop for {0}: {1}", strServerName, DateTime.Now.ToLocalTime().ToString("MMddyyy HH:mm:ss"));

            if (bPingSuccess)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        static void funcPrintParameterWarning()
        {
            Console.WriteLine("A parameter must be specified to run ServiceAccountFinder.");
            Console.WriteLine("Run ServiceAccountFinder -? to get the parameter syntax.");
        }

        static void funcPrintParameterSyntax()
        {
            Console.WriteLine("ServiceAccountFinder");
            Console.WriteLine();
            Console.WriteLine("Parameter syntax:");
            Console.WriteLine();
            Console.WriteLine("Use the following for the first parameter:");
            Console.WriteLine("-run                required parameter");
            Console.WriteLine();
            Console.WriteLine("Examples:");
            Console.WriteLine("ServiceAccountFinder -run");
        } // funcPrintParameterSyntax

        static void funcLogToEventLog(string strAppName, string strEventMsg, int intEventType)
        {
            string sLog;

            sLog = "Application";

            if (!EventLog.SourceExists(strAppName))
                EventLog.CreateEventSource(strAppName, sLog);

            //EventLog.WriteEntry(strAppName, strEventMsg);
            EventLog.WriteEntry(strAppName, strEventMsg, EventLogEntryType.Information, intEventType);

        } // LogToEventLog

        static CMDArguments funcParseCmdArguments(string[] cmdargs)
        {
            CMDArguments objCMDArguments = new CMDArguments();

            try
            {
                objCMDArguments.bParseCmdArguments = false;

                if (cmdargs[0] == "-run" & cmdargs.Length == 1)
                {
                    objCMDArguments.bParseCmdArguments = true;
                }
                else
                {
                    objCMDArguments.bParseCmdArguments = false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                objCMDArguments.bParseCmdArguments = false;
            }

            return objCMDArguments;
        }

        static void funcProgramExecution(CMDArguments objCMDArguments2)
        {
            try
            {
                funcLogToEventLog("ServiceAccountFinder", "ServiceAccountFinder started successfully.", 1301);

                string strQueryFilter = "(&(&(&(sAMAccountType=805306369)(objectCategory=computer)(|(operatingSystem=Windows Server 2008*)(operatingSystem=Windows Server 2003*)(operatingSystem=Windows 2000 Server*)(operatingSystem=Windows NT*)(operatingSystem=*2008*)))))";

                List<string> lstExcludeServiceName = new List<string>();

                lstExcludeServiceName.Add("LocalSystem");
                lstExcludeServiceName.Add("NT AUTHORITY\\LocalService");
                lstExcludeServiceName.Add("NT Authority\\LocalService");
                lstExcludeServiceName.Add("NT AUTHORITY\\LOCALSERVICE");
                lstExcludeServiceName.Add("NT AUTHORITY\\LOCAL SERVICE");
                lstExcludeServiceName.Add("NT AUTHORITY\\NetworkService");
                lstExcludeServiceName.Add("NT Authority\\NetworkService");
                lstExcludeServiceName.Add("NT AUTHORITY\\NETWORKSERVICE");
                lstExcludeServiceName.Add("NT AUTHORITY\\NETWORK SERVICE");
                lstExcludeServiceName.Add("localSystem");
                //lstExcludeServiceName.Add("");

                // [Comment] Get local domain context
                string rootDSE;

                System.DirectoryServices.DirectorySearcher objrootDSESearcher = new System.DirectoryServices.DirectorySearcher();
                rootDSE = objrootDSESearcher.SearchRoot.Path;
                // [DebugLine]Console.WriteLine(rootDSE);

                // [Comment] Construct DirectorySearcher object using rootDSE string
                System.DirectoryServices.DirectoryEntry objrootDSEentry = new System.DirectoryServices.DirectoryEntry(rootDSE);
                System.DirectoryServices.DirectorySearcher objComputerObjectSearcher = new System.DirectoryServices.DirectorySearcher(objrootDSEentry);
                // [DebugLine]Console.WriteLine(objComputerObjectSearcher.SearchRoot.Path);

                // [Comment] Add filter to DirectorySearcher object
                objComputerObjectSearcher.Filter = (strQueryFilter);

                // [Comment] Execute query, return results, display name and path values
                System.DirectoryServices.SearchResultCollection objComputerResults = objComputerObjectSearcher.FindAll();
                // [DebugLine] Console.WriteLine(objComputerResults.Count.ToString());

                TextWriter twOutputLog = funcOpenOutputLog();
                string strOutputMsg = "";

                foreach(SearchResult sr in objComputerResults)
                {
                    DirectoryEntry newDE = sr.GetDirectoryEntry();
                    // [DebugLine] Console.WriteLine(newDE.Name.Substring(3));
                
                    ManagementObjectCollection oQueryCollection = null;
                    string strHostName = "";

                    strHostName = newDE.Name.Substring(3);

                    if (funcPingServer(strHostName))
                    {
                        //**********************************
                        //Begin-Win32_Service
                        //**********************************
                        //Get the query results for Win32_Service
                        oQueryCollection = null;
                        oQueryCollection = funcSysQueryData("select * from Win32_Service", strHostName);

                        // [DebugLine] Console.WriteLine("{0} \t Service Count: {1}", strHostName, oQueryCollection.Count.ToString());
                        strOutputMsg = String.Format("{0} \t Service Count: {1}", strHostName, oQueryCollection.Count.ToString());
                        funcWriteToOutputLog(twOutputLog, strOutputMsg);

                        foreach (ManagementObject oReturn in oQueryCollection)
                        {
                            // "Caption","Description","PathName","Status","State","StartMode","StartName"

                            //string[] strElementBag = new string[] { "Caption", "Description", "PathName", "Status", "State", "StartMode", "StartName" };
                            //foreach (string strElement in strElementBag)
                            //{
                            //    string strElementTemp = strElement.ToLower(new CultureInfo("en-US", false));
                            //    try
                            //    {
                            //        Console.WriteLine("\"" + strElementTemp + "\"" + " : " + "\"" + oReturn[strElement].ToString().Trim() + "\"");
                            //    }
                            //    catch
                            //    {
                            //        Console.WriteLine("\"" + strElementTemp + "\"" + " : " + "\"" + "<na>" + "\"");
                            //    }
                            //}
                            if (!lstExcludeServiceName.Contains(oReturn.Properties["StartName"].Value.ToString()))
                            {
                                Console.WriteLine("{0} \t {1} \t {2}", strHostName, oReturn.Properties["Caption"].Value.ToString(), oReturn.Properties["StartName"].Value.ToString());
                            }

                        }
                        //**********************************
                        //End-Win32_Service
                        //**********************************
                    }
                    else
                    {
                        strOutputMsg = String.Format("{0} \t {1}", strHostName, ".");
                        funcWriteToOutputLog(twOutputLog, strOutputMsg);
                    }
                }

                funcCloseOutputLog(twOutputLog);
                
                funcLogToEventLog("ServiceAccountFinder", "ServiceAccountFinder stopped.", 1302);

            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcGetFuncCatchCode(string strFunctionName, Exception currentex)
        {
            string strCatchCode = "";

            Dictionary<string, string> dCatchTable = new Dictionary<string, string>();
            dCatchTable.Add("funcGetFuncCatchCode", "f0");
            dCatchTable.Add("funcPrintParameterWarning", "f2");
            dCatchTable.Add("funcPrintParameterSyntax", "f3");
            dCatchTable.Add("funcParseCmdArguments", "f4");
            dCatchTable.Add("funcProgramExecution", "f5");
            dCatchTable.Add("funcCreateDSSearcher", "f7");
            dCatchTable.Add("funcCreatePrincipalContext", "f8");
            dCatchTable.Add("funcCheckNameExclusion", "f9");
            dCatchTable.Add("funcMoveDisabledAccounts", "f10");
            dCatchTable.Add("funcFindAccountsToDisable", "f11");
            dCatchTable.Add("funcCheckLastLogin", "f12");
            dCatchTable.Add("funcRemoveUserFromGroup", "f13");
            dCatchTable.Add("funcToEventLog", "f14");
            dCatchTable.Add("funcCheckForFile", "f15");
            dCatchTable.Add("funcCheckForOU", "f16");
            dCatchTable.Add("funcWriteToErrorLog", "f17");

            if (dCatchTable.ContainsKey(strFunctionName))
            {
                strCatchCode = "err" + dCatchTable[strFunctionName] + ": ";
            }

            //[DebugLine] Console.WriteLine(strCatchCode + currentex.GetType().ToString());
            //[DebugLine] Console.WriteLine(strCatchCode + currentex.Message);

            funcWriteToErrorLog(strCatchCode + currentex.GetType().ToString());
            funcWriteToErrorLog(strCatchCode + currentex.Message);

        }

        static void funcWriteToErrorLog(string strErrorMessage)
        {
            try
            {
                FileStream newFileStream = new FileStream("Err-ServiceAccountFinder.log", FileMode.Append, FileAccess.Write);
                TextWriter twErrorLog = new StreamWriter(newFileStream);

                DateTime dtNow = DateTime.Now;

                string dtFormat = "MMddyyyy HH:mm:ss";

                twErrorLog.WriteLine("{0} \t {1}", dtNow.ToString(dtFormat), strErrorMessage);

                twErrorLog.Close();
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }

        }

        static bool funcCheckForOU(string strOUPath)
        {
            try
            {
                string strDEPath = "";

                if (!strOUPath.Contains("LDAP://"))
                {
                    strDEPath = "LDAP://" + strOUPath;
                }
                else
                {
                    strDEPath = strOUPath;
                }

                if (DirectoryEntry.Exists(strDEPath))
                {
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static bool funcCheckForFile(string strInputFileName)
        {
            try
            {
                if (System.IO.File.Exists(strInputFileName))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return false;
            }
        }

        static TextWriter funcOpenOutputLog()
        {
            try
            {
                DateTime dtNow = DateTime.Now;

                string dtFormat2 = "MMddyyyy"; // for log file directory creation

                string strPath = Directory.GetCurrentDirectory();

                string strLogFileName = strPath + "\\ServiceAccountFinder" + dtNow.ToString(dtFormat2) + ".log";

                FileStream newFileStream = new FileStream(strLogFileName, FileMode.Append, FileAccess.Write);
                TextWriter twOuputLog = new StreamWriter(newFileStream);

                return twOuputLog;
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
                return null;
            }

        }

        static void funcWriteToOutputLog(TextWriter twCurrent, string strOutputMessage)
        {
            try
            {
                DateTime dtNow = DateTime.Now;

                string dtFormat = "MMddyyyy HH:mm:ss";

                twCurrent.WriteLine("{0} \t {1}", dtNow.ToString(dtFormat), strOutputMessage);
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void funcCloseOutputLog(TextWriter twCurrent)
        {
            try
            {
                twCurrent.Close();
            }
            catch (Exception ex)
            {
                MethodBase mb1 = MethodBase.GetCurrentMethod();
                funcGetFuncCatchCode(mb1.Name, ex);
            }
        }

        static void Main(string[] args)
        {
            try
            {
                if (args.Length == 0)
                {
                    funcPrintParameterWarning();
                }
                else
                {
                    if (args[0] == "-?")
                    {
                        funcPrintParameterSyntax();
                    }
                    else
                    {
                        string[] arrArgs = args;
                        CMDArguments objArgumentsProcessed = funcParseCmdArguments(arrArgs);

                        if (objArgumentsProcessed.bParseCmdArguments)
                        {
                            funcProgramExecution(objArgumentsProcessed);
                        }
                        else
                        {
                            funcPrintParameterWarning();
                        } // check objArgumentsProcessed.bParseCmdArguments
                    } // check args[0] = "-?"
                } // check args.Length == 0
            }
            catch (Exception ex)
            {
                Console.WriteLine("errm0: {0}", ex.Message);
            }
        }
    }
}
