using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OPCAutomation;
using Services;
namespace CommunicationServers.OPC
{
    
    public class OpcHelper : IDisposable
    {
        private string strHostIP;
        private string strHostName;
        private OPCServer opcServer;
        private OPCGroups opcGroups;
        private OPCGroup opcGroup;
        private List<int> itemHandleClient = new List<int>();
        private List<int> itemHandleServer = new List<int>();
        private List<string> itemNames = new List<string>();
        private List<string> oPCServerNames = new List<string>();
        private OPCItems opcItems;
        private OPCItem opcItem;
        private Dictionary<string, string> itemValues = new Dictionary<string, string>();
        public bool Connected = false;
        public List<OPCDataItem> OpcDataItems = new List<OPCDataItem>();
        public event DelegateDataChange DataChangeEvent;
        private int opcServerState;
        public List<string> OPCServerNames
        {
            get
            {
                return oPCServerNames;
            }
        }

        public int OpcServerState
        {
            get
            {
                return opcServer.ServerState;
            }
        }

        public OpcHelper()
        {

        }
        public OpcHelper(string strHostIP,string strHostName,int UpdateRate)
        {
            //this.strHostIP = strHostIP;
            //this.strHostName = strHostName;
            //if (!CreateServer())
            //    return;
            //if (!ConnectServer(strHostIP, strHostName))
            //    return;
            //Connected = true;
            //opcGroups = opcServer.OPCGroups;
            //opcGroup = opcGroups.Add("GQOPCGROUP");
        }

        /// <summary>
        /// 创建服务
        /// </summary>
        /// <returns></returns>
        public bool CreateServer(string IP)
        {
            try
            {
                opcServer = new OPCServer();
                object serverList = opcServer.GetOPCServers(IP);
                foreach (string item in (Array)serverList)
                {
                    oPCServerNames.Add(item);
                }
            }
            catch(Exception ex)
            {
                SimpleLogHelper.Instance.WriteLog(LogType.Error, ex);
                return false;
            }
            return true;
        }

        /// <summary>
        /// 连接到服务器
        /// </summary>
        /// <param name="strHostIP"></param>
        /// <param name="strHostName"></param>
        /// <returns></returns>
        public bool ConnectServer(string strHostIP,string strHostName)
        {
            try
            {
                opcServer.Connect(strHostName, strHostIP);
            }
            catch(Exception ex)
            {
                SimpleLogHelper.Instance.WriteLog(LogType.Error, ex);
                return false;
            }
            return true;
        }

        public bool CreateNewGroup(string groupName,int updateRate)
        {
            try
            {
                opcGroups = opcServer.OPCGroups;
                opcGroups.DefaultGroupIsActive = true;
                opcGroup = opcGroups.Add(groupName);
                opcGroup.IsActive = true;
                opcGroup.DeadBand = 0;
                opcGroup.UpdateRate = updateRate;
                opcGroup.IsSubscribed = true;
                opcGroup.DataChange += new DIOPCGroupEvent_DataChangeEventHandler(OpcGroup_DataChange);
                opcItems = opcGroup.OPCItems;
            }
            catch
            {
                return false;
            }
            return true;
        }

        private void OpcGroup_DataChange(int TransactionID, int NumItems, ref Array ClientHandles, ref Array ItemValues, ref Array Qualities, ref Array TimeStamps)
        {
            for (int i = 1; i <= NumItems; i++)
            {
                object handleValue = ClientHandles.GetValue(i);
                int index = OpcDataItems.FindIndex(m => m.ItemHandle.ToString() == handleValue.ToString());
                if (index != -1)
                {
                    OpcDataItems[index].ItemValue = ItemValues.GetValue(i);
                    OpcDataItems[index].Quality = Qualities.GetValue(i);
                    OpcDataItems[index].TimeStamp = TimeStamps.GetValue(i);
                }
            }
            if (DataChangeEvent == null) return;
            DataChangeEvent(OpcDataItems);

        }

        public bool AddItem(string itemName)
        {
            try
            {
                opcItem = opcItems.AddItem(itemName, opcItems.Count + 1);
                Console.WriteLine(opcItems.Count);
                OpcDataItems.Add(new OPCDataItem { ItemName = itemName, ItemHandle = opcItems.Count });
            }
            catch
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 展开OPC服务器的节点
        /// </summary>
        /// <param name="opcServer">OPC服务器</param>
        /// <returns>返回展开后的节点数据</returns>
        public List<string> RecurBrowse()
        {
            OPCBrowser opcBrowser = opcServer.CreateBrowser();
            //展开分支
            opcBrowser.ShowBranches();
            //展开叶子
            opcBrowser.ShowLeafs(true);
            List<string> browserNames = new List<string>();
            for (int i = 1; i <= opcBrowser.Count; i++)
            {
                browserNames.Add(opcBrowser.Item(i));
            }
            return browserNames;
        }


        //----------------------下面的目前没用到
        private void SetGroupProperty(OPCGroup opcGroup,int updateRate)
        {
            opcGroup.IsActive = true;
            opcGroup.DeadBand = 0;
            opcGroup.UpdateRate = updateRate;
            opcGroup.IsSubscribed = true;
        }

        public bool Contains(string itemNameContains)
        {
            foreach(string key in itemValues.Keys)
            {
                if (key == itemNameContains)
                    return true;
            }
            return false;
        }

        public void AddItems(string[] itemNamesAdded)
        {
            for(int i = 0;i<itemNamesAdded.Length;i++)
            {
                this.itemNames.Add(itemNamesAdded[i]);
                itemValues.Add(itemNamesAdded[i], "");
            }
            for(int i = 0; i <itemNamesAdded.Length;i++)
            {
                itemHandleClient.Add(itemHandleClient.Count != 0 ? itemHandleClient[itemHandleClient.Count - 1] + 1 : 1);
                opcItem = opcItems.AddItem(itemNamesAdded[i], itemHandleClient[itemHandleClient.Count - 1]);
                itemHandleServer.Add(opcItem.ServerHandle);
            }
        }

        public string[] GetItemsValues(string[] getValuesItemNames)
        {
            string[] getedValues = new string[getValuesItemNames.Length];
            for(int i = 0;i<getValuesItemNames.Length;i++)
            {
                if (Contains(getValuesItemNames[i]))
                    itemValues.TryGetValue(getValuesItemNames[i], out getedValues[i]);
            }
            return getedValues;
        }

        private void opcGroup_DataChange(int TransactionID, int NumItems,ref Array ClientHandles , ref Array ItemValues, ref Array Qualities,ref Array TimeStamps)
        {
            for(int i = 1; i<= NumItems;i++)
            {
                itemValues[itemNames[Convert.ToInt32(ClientHandles.GetValue(i)) - 1]] = ItemValues.GetValue(i).ToString();
            }
        }

        public void  AsyncWrite(string[] writeItemsNames,string[] writeItemValues)
        {
            OPCItem[] bItem = new OPCItem[writeItemsNames.Length];
            for (int i = 0; i < writeItemsNames.Length; i++)
            {
                for (int j = 0; j < itemNames.Count; j++)
                {
                    if(itemNames[j] == writeItemsNames[i])
                    {
                        bItem[i] = opcItems.GetOPCItem(itemHandleServer[j]);
                        break;
                    }
                }
            }
            int[] temp = new int[writeItemsNames.Length + 1];
            temp[0] = 0;
            for (int i = 1; i < writeItemsNames.Length + 1; i++)
            {
                temp[i] = bItem[i - 1].ServerHandle;
            }
            Array serverHandles = (Array)temp;
            object[] valueTemp = new object[writeItemsNames.Length + 1];
            valueTemp[0] = "";
            for(int i =1;i<writeItemsNames.Length +1;i++)
            {
                valueTemp[i] = writeItemsNames[i - 1];
                Array values = (Array)valueTemp;
                Array Errors;
                int cancelID;
                opcGroup.AsyncWrite(writeItemsNames.Length, ref serverHandles, ref values, out Errors, 2009, out cancelID);
                GC.Collect();
            }
        }

        public void SyncWrite(string[] writeItemsNames, string[] writeItemValues)
        {
            OPCItem[] bItem = new OPCItem[writeItemsNames.Length];
            for (int i = 0; i < writeItemsNames.Length; i++)
            {
                for (int j = 0; j < itemNames.Count; j++)
                {
                    if (itemNames[j] == writeItemsNames[i])
                    {
                        bItem[i] = opcItems.GetOPCItem(itemHandleServer[j]);
                    }
                }
            }
            int[] temp = new int[writeItemsNames.Length + 1];
            temp[0] = 0;
            for (int i = 1; i < writeItemsNames.Length + 1; i++)
            {
                temp[i] = bItem[i - 1].ServerHandle;
            }
            Array serverHandles = (Array)temp;
            object[] valueTemp = new object[writeItemsNames.Length + 1];
            valueTemp[0] = "";
            for (int i = 1; i < writeItemsNames.Length + 1; i++)
            {
                valueTemp[i] = writeItemsNames[i - 1];
                Array values = (Array)valueTemp;
                Array Errors;
                opcGroup.SyncWrite(writeItemsNames.Length, ref serverHandles, ref values, out Errors);
                GC.Collect();
            }
        }

        void opcGroup_AsyncWriteComplete(int TransactionID,int NumItems,ref Array ClientHandles,ref Array Errors)
        {
            throw new NotImplementedException();
        }

        public void Dispose()
        {
            if(opcGroup != null)
            {
                opcGroup.DataChange -= new DIOPCGroupEvent_DataChangeEventHandler(opcGroup_DataChange);
                opcGroup.AsyncWriteComplete -= new DIOPCGroupEvent_AsyncWriteCompleteEventHandler(opcGroup_AsyncWriteComplete);
            }
            if(opcServer != null)
            {
                opcServer.Disconnect();
                opcServer = null;
            }
            Connected = false;
        }
    }
}
