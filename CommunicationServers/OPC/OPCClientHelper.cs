using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OPCAutomation;
using Services;
using System.Threading;

namespace CommunicationServers.OPC
{
    public delegate void DelegateDataChange(List<OPCDataItem> OpcDataItems);

    public class OPCClientHelper
    {
        OPCGroups KepGroups;
        OPCGroup KepGroup;
        OPCItems KepItems;
        OPCItem KepItem;
        public Array KeyItemServerHandles;
        List<int> KepItemServerHandles = new List<int>();
        public List<OPCDataItem> OpcDataItems = new List<OPCDataItem>();
        public List<string> Items;
        public event DelegateDataChange DataChangeEvent;
        public OPCClientHelper()
        {
            OpcDataItems = new List<OPCDataItem>();
            KepItemServerHandles.Insert(0, 0); //加一项作为首项
        }

        /// <summary>
        /// 获取OPC服务名称
        /// </summary>
        /// <param name="serverIP">服务器IP</param>
        /// <returns></returns>
        public List<string> GetOPCServerNames(string serverIP)
        {
            OPCServer opcServer = new OPCServer();
            List<string> lst = new List<string>();
            object serverList = opcServer.GetOPCServers(serverIP);
            foreach (string item in (Array)serverList)
            {
                lst.Add(item);
            }
            return lst;
        }

        /// <summary>
        /// 连接OPCServer
        /// </summary>
        /// <param name="serverName">Server名称</param>
        /// <param name="serverIP">服务器IP</param>
        /// <returns></returns>
        public OPCServer ConnectToServer(string serverName, string serverIP)
        {
            OPCServer opcServer = new OPCServer();
            try
            {
                opcServer.Connect(serverName, serverIP);
                if (opcServer.ServerState != (int)OPCServerState.OPCRunning)
                {
                    opcServer.Disconnect();
                    return null;
                }
            }
            catch
            {
                opcServer.Disconnect();
                return null;
            }
            return opcServer;
        }

        /// <summary>
        /// 获取OPC服务器的相关信息
        /// </summary>
        /// <param name="opcServer">OPC服务器对象</param>
        /// <returns>OPC服务器信息</returns>
        public OPCServerInfo GetServerInfo(OPCServer opcServer)
        {
            OPCServerInfo serverInfo = new OPCServerInfo();
            serverInfo.StartTime = opcServer.StartTime;
            serverInfo.ServerVersion = opcServer.MajorVersion.ToString() + "." + opcServer.MinorVersion.ToString() + "." + opcServer.BuildNumber.ToString();
            return serverInfo;
        }

        /// <summary>
        /// 展开OPC服务器的节点
        /// </summary>
        /// <param name="opcServer">OPC服务器</param>
        /// <returns>返回展开后的节点数据</returns>

        public OPCBrowser RecurBrowse(OPCServer opcServer)
        {
            OPCBrowser opcBrowser = opcServer.CreateBrowser();
            //展开分支
            opcBrowser.ShowBranches();
            //展开叶子
            opcBrowser.ShowLeafs(true);
            return opcBrowser;
        }

        public  OPCGroup CreateGroup(OPCServer opcServer, OPCItems opcItems, string opcGroupName, GroupProperty groupProperty)
        {
            try
            {
                OPCGroup opcGroup = opcServer.OPCGroups.Add(opcGroupName);
                //opcServer.OPCGroups.DefaultGroupIsActive = groupProperty.DefaultGroupIsActive;
                //opcServer.OPCGroups.DefaultGroupDeadband = groupProperty.DefaultGroupDeadband;
                //opcGroup.UpdateRate = groupProperty.UpdateRate;
                //opcGroup.IsActive = groupProperty.IsActive;
                //opcGroup.IsSubscribed = groupProperty.IsSubscribed;
                opcServer.OPCGroups.DefaultGroupIsActive = true;
                opcServer.OPCGroups.DefaultGroupDeadband = 0;
                opcGroup.UpdateRate = 1000;
                opcGroup.DataChange += OpcGroup_DataChange;
                //opcGroup.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(OpcGroupDataChange);
                //opcGroup.AsyncWriteComplete += new DIOPCGroupEvent_AsyncWriteCompleteEventHandler(KepGroup_AsyncWriteComplete);
                //opcItems = opcGroup.OPCItems;
                return opcGroup;
            }
            catch (Exception ex)
            {
                SimpleLogHelper.Instance.WriteLog(LogType.Error, ex);
                return null;
            }
        }

        public OPCGroup CreateGroup(OPCServer opcServer,string opcGroupName)
        {
            try
            {
                //opcServer.OPCGroups.DefaultGroupIsActive = true;
                //opcServer.OPCGroups.DefaultGroupDeadband = 0;
                //opcServer.OPCGroups.DefaultGroupUpdateRate = 1000;
                KepGroups = opcServer.OPCGroups;
                KepGroup = KepGroups.Add(opcGroupName);
                KepGroup.IsActive = true;
                KepGroup.DeadBand = 0;
                KepGroup.UpdateRate = 1000;
                KepGroup.IsSubscribed = true;
                //KepGroup.AsyncReadComplete += KepGroup_AsyncReadComplete;
                KepGroup.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(OpcGroup_DataChange);
                KepGroup.IsActive = true;
                KepGroup.IsSubscribed = true;
                //AsyncRead();
                return KepGroup;
            }
            catch (Exception ex)
            {
                SimpleLogHelper.Instance.WriteLog(LogType.Error, ex);
                return null;
            }
        }

        public void AddItem(string itemName)
        {
            try
            {
                KepItems = KepGroup.OPCItems;
                KepItem = KepItems.AddItem(itemName, KepItemServerHandles.Count);
                OpcDataItems.Add(new OPCDataItem { ItemName = itemName, ItemHandle = KepItemServerHandles.Count });
                KepItemServerHandles.Add(KepItem.ServerHandle);
                KeyItemServerHandles = KepItemServerHandles.ToArray();
            }
            catch(Exception ex)
            {
                SimpleLogHelper.Instance.WriteLog(LogType.Error, ex);
            }
        }

        public void AsyncRead()
        {
            Array Errors;
            int cancelID;
            try
            {
                KepGroup.AsyncRead(KeyItemServerHandles.Length, ref KeyItemServerHandles, out Errors, 2009, out cancelID);
            }
            catch (Exception exp1)
            {
            }
        }

        private void KepGroup_AsyncReadComplete(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps, ref Array Errors)
        {
            for (int i = 1; i <= NumItems; i++)
            {
                Console.WriteLine(ClientHandles.GetValue(i));
            }
        }

        private void OpcGroup_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            for (int i = 1; i <= NumItems; i++)
            {
                object handleValue = ClientHandles.GetValue(i);
                int index = OpcDataItems.FindIndex(m => m.ItemHandle.ToString() == handleValue.ToString());
                if(index != -1)
                {
                    OpcDataItems[index].ItemValue = ItemValues.GetValue(i);
                    OpcDataItems[index].Quality = Qualities.GetValue(i);
                    OpcDataItems[index].TimeStamp = TimeStamps.GetValue(i);
                }
            }
            if (DataChangeEvent == null) return;
            DataChangeEvent(OpcDataItems);
            
        }
    }

    public class OPCServerInfo
    {
        public DateTime StartTime { get; set; }
        public string ServerVersion { get; set; }
    }

    public class OPCDataItem
    {
        public object ItemName { get; set; }
        public object ItemHandle { get; set; }
        public object ItemValue { get; set; }
        public object Quality { get; set; }
        public object TimeStamp { get; set; }

    }

    public class GroupProperty
    {
        public bool DefaultGroupIsActive { get; set; }
        public float DefaultGroupDeadband { get; set; }
        public int UpdateRate { get; set; }
        public bool IsActive { get; set; }
        public bool IsSubscribed { get; set; }

        public GroupProperty()
        {
            DefaultGroupIsActive = true;
            DefaultGroupDeadband = 0;
            UpdateRate = 250;
            IsActive = true;
            IsSubscribed = true;
        }
    }
}

/*
 opc连接步骤：
   1、
   2、
   3、
   4、
   5、
    
 */