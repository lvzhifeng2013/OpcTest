using Opc;
using Opc.Da;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpcTest
{
    class OPCHelper
    {
        #region OPC相关的全局变量
        public static Opc.Da.Server OpcGroupClassm_server = null;//定义数据存取服务器
        public static Opc.Da.Subscription OpcGroupClasssubscription = null;//定义组对象（订阅者）
        public static Opc.Da.SubscriptionState OpcGroupClassstate = null;//定义组（订阅者）状态，相当于OPC规范中组的参数      
        public static Opc.IDiscovery OpcGroupClassm_discovery = new OpcCom.ServerEnumerator();//定义枚举基于COM服务器的接口，用来搜索所有的此类服务器     
        public static string OpcServerName;//OPC服务器名称
        #endregion
        Dictionary<string, Test> dicTest;//地址字典(Key：AddressName；值：地址) 
        public void OnDataChange(object subscriptionHandle, object requestHandle, ItemValueResult[] values)
        {
            foreach (ItemValueResult item in values)
            {
                if (item.Quality.GetCode() == 192)//
                {
                    var address = dicTest[item.ItemName];
                    address.Value = item.Value.ToString();
                    switch (address.Description)
                    {
                       


                    }
                }
            }
        }

        #region OPC相关函数方法
        /// <summary>
        /// OPC连接
        /// </summary>
        /// <param name="barcode"></param>
        public void OPC_connect()
        {

            #region 连接OPCSever
            //if (!OpcGroupClassm_server.IsConnected) { return; }
            //获取本地的IP地址
            var addr = "192.168.3.89";
            OpcServerName = addr + ".KEPware.KEPServerEx.V6";
            Opc.Server[] servers = OpcGroupClassm_discovery.GetAvailableServers(Specification.COM_DA_20, addr, null);//查询服务器
            if (servers != null)
            {
                foreach (Opc.Da.Server server in servers)
                {
                    //server即为需要连接的OPC数据存取服务器
                    if (String.Compare(server.Name, OpcServerName, true) == 0)//true表示忽略大小写
                    {
                        OpcGroupClassm_server = server;//建立连接
                        break;
                    }
                }
            }
            //连接服务器
            if (OpcGroupClassm_server != null) OpcGroupClassm_server.Connect();//非空连接服务器
            #endregion

            #region 往OPCSever中添加Group
            OpcGroupClassstate = new Opc.Da.SubscriptionState();//组（订阅者）状态，相当于OPC规范中组的参数
            OpcGroupClassstate.Name = "OpcGroupClasss";//组名
            OpcGroupClassstate.ServerHandle = null;//服务器给该组分配的句柄。
            OpcGroupClassstate.ClientHandle = Guid.NewGuid().ToString();//客户端给该组分配的句柄
            OpcGroupClassstate.Active = true;//激活该组
            OpcGroupClassstate.UpdateRate = 100;//刷新频率为1秒
            OpcGroupClassstate.Deadband = 0;// 死区值，设为0时，服务器端该组内任何数据变化都通知组
            OpcGroupClassstate.Locale = null;//不设置地区值
            OpcGroupClasssubscription = (Opc.Da.Subscription)OpcGroupClassm_server.CreateSubscription(OpcGroupClassstate);//添加组
                                                                                                                          //添加监控地址
            #endregion

            #region  往Group添加 Items
            try
            {
                List<Item> items = new List<Item>();
                foreach (var adress in dicTest)
                {
                    var item = new Item()
                    {
                        ClientHandle = Guid.NewGuid().ToString(),//客户端给该数据项分配的句柄。
                        ItemName = adress.Key //该数据项在服务器中的名字
                    };
                    items.Add(item);
                }
                OpcGroupClasssubscription.AddItems(items.ToArray());
            }
            catch (Exception ex)
            {
                MessageBox.Show("初始化监控地址失败：" + ex.Message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }

            OpcGroupClasssubscription.DataChanged += new DataChangedEventHandler(OnDataChange);
            //Thread.Sleep(1000);
            #endregion
        }

        /// <summary>
        /// 断开OPC连接
        /// </summary>
        public void OPC_disconnect()
        {
            OpcGroupClasssubscription.DataChanged -= new Opc.Da.DataChangedEventHandler(this.OnDataChange);//取消回调事件
            OpcGroupClasssubscription.RemoveItems(OpcGroupClasssubscription.Items);//移除组内item
            //结束：释放各资源
            OpcGroupClassm_server.CancelSubscription(OpcGroupClasssubscription);//m_server前文已说明，通知服务器要求删除组。        
            OpcGroupClasssubscription.Dispose();//强制.NET资源回收站回收该subscription的所有资源。         
            OpcGroupClassm_server.Disconnect();//断开服务器连接
            //不相关的
        }

        /// <summary>
        /// 根据ItemName找到OPC中Item项
        /// </summary>
        /// <param name="itemName"></param>
        /// <returns></returns>
        public Item FindOpcItem(string itemName)
        {
            Item opcItem = null;

            foreach (var item in OpcGroupClasssubscription.Items)
            {
                if (itemName == item.ItemName)
                {
                    opcItem = item;
                    break;
                }
            }
            return opcItem;
        }

        /// <summary>
        /// 根据ItemName数组找到OPC中Item项的数组
        /// </summary>
        /// <param name="itemName"></param>
        /// <returns></returns>
        public List<Item> FindOpcItems(List<string> itemNames)
        {
            List<Item> opcItems = new List<Item>();
            foreach (var itemName in itemNames)
            {
                foreach (var item in OpcGroupClasssubscription.Items)
                {
                    if (itemName == item.ItemName)
                    {
                        opcItems.Add(item);
                        break;
                    }
                }
            }
            return opcItems;
        }

        /// <summary>
        /// OPC同步写一个数据
        /// </summary>
        /// <param name="opcitem">地址Name</param>
        /// <param name="WriteVaule">待写的地址Value</param>
        /// <returns>是否写成功</returns>
        public bool SynWriteOpcItem(string itemName, string writeValue)
        {
            try
            {
                ItemValue[] itemvalues = new ItemValue[1];
                foreach (var item in OpcGroupClasssubscription.Items)
                {
                    if (itemName == item.ItemName)
                    {
                        itemvalues[0] = new ItemValue((ItemIdentifier)item);
                        itemvalues[0].Value = writeValue;
                        break;
                    }
                }
                OpcGroupClasssubscription.Write(itemvalues);
                return true;
            }
            catch (Exception ex)
            {
                //ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
                //log.Fatal("写心跳信号失败！", ex);
                return false;
            }
        }

        /// <summary>
        /// OPC同步写多个数据
        /// </summary>
        /// <param name="writeItems">需要些数据的地址和字符串值对应的字典（Key：地址Name；Value：待写的地址Value）</param>
        /// <returns>是否写成功</returns>
        public bool SynWriteOpcItems(Dictionary<string, string> writeItems)
        {
            try
            {
                var itemValues = new List<ItemValue>();
                foreach (var writeItem in writeItems)
                {
                    foreach (var item in OpcGroupClasssubscription.Items)
                    {
                        if (writeItem.Key == item.ItemName)
                        {
                            var itemValue = new ItemValue((ItemIdentifier)item);
                            itemValue.Value = writeItem.Value;
                            itemValues.Add(itemValue);
                            break;
                        }
                    }
                }
                OpcGroupClasssubscription.Write(itemValues.ToArray());
                return true;
            }
            catch (Exception ex)
            {
                //ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
                //log.Fatal("OPC写数据失败！", ex);
                return false;
            }
        }

        /// <summary>
        /// OPC同步读一个数据
        /// </summary>
        /// <param name="itemName"></param>
        /// <returns></returns>
        public string SynReadOpcItem(string itemName)
        {
            try
            {
                var item = FindOpcItem(itemName);
                Item[] readItems = { item };
                ItemValueResult[] itemValues = OpcGroupClasssubscription.Read(readItems);
                ItemValueResult itemValue = itemValues[0];
                if (itemValue.Quality == Quality.Bad)
                    return null;
                else
                    return itemValue.Value.ToString();
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// OPC同步读多个数据
        /// </summary>
        /// <param name="itemNames"></param>
        /// <returns></returns>
        public Dictionary<string, string> SynReadOpcItems(List<string> itemNames)
        {
            try
            {
                var itemValues = new Dictionary<string, string>();

                var items = FindOpcItems(itemNames);
                Item[] readItems = items.ToArray();
                ItemValueResult[] itemValueResults = OpcGroupClasssubscription.Read(readItems);

                foreach (var result in itemValueResults)
                {
                    itemValues.Add(result.ItemName, result.Value.ToString());
                }

                return itemValues;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public List<string> SynReadOpcItems2(List<string> itemNames)//同步读多个数据生成动态数组
        {
            try
            {
                List<string> itemValues = new List<string>();
                var items = FindOpcItems(itemNames);
                Item[] readItems = items.ToArray();
                ItemValueResult[] itemValueResults = OpcGroupClasssubscription.Read(readItems);

                foreach (var result in itemValueResults)
                {
                    itemValues.Add(result.Value.ToString());
                }

                return itemValues;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        #endregion

    }
}
