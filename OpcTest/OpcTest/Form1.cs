using Opc;
using Opc.Da;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace OpcTest
{
    public partial class Form1 : Form
    {
        #region OPC相关的全局变量
        public static Opc.Da.Server OpcGroupClassm_server = null;//定义数据存取服务器
        public static Opc.Da.Subscription OpcGroupClasssubscription = null;//定义组对象（订阅者）
        public static Opc.Da.SubscriptionState OpcGroupClassstate = null;//定义组（订阅者）状态，相当于OPC规范中组的参数      
        public static Opc.IDiscovery OpcGroupClassm_discovery = new OpcCom.ServerEnumerator();//定义枚举基于COM服务器的接口，用来搜索所有的此类服务器     
        public static string OpcServerName;//OPC服务器名称
        #endregion
        #region 全局变量
        Dictionary<string, Test> dicTest;//地址字典(Key：AddressName；值：地址)     
        string KepIP;
        #endregion
        public Form1()
        {
            InitializeComponent();
        }

        #region 窗体事件
        private void Form1_Load(object sender, EventArgs e)
        {

            try
            {
                string path = Application.StartupPath + @"\Config\KepConfig.xml";//配置文件"LED1.bmp";
                KepIP = getKepIP(path);
                DataTable dt = DbHelperSQL.OpenTable("select * from opc ");
                dicTest = new Dictionary<string, Test>();

                foreach (DataRow row in dt.Rows)
                {
                    Test test = new Test(row["id"].ToString(), row["address_content"].ToString(), row["description"].ToString());
                    dicTest.Add(test.Id, test);
                }
                dataGridView1.DataSource = dt;
                OPC_connect();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.Rows.Count - 1;
        }
        private void dataGridView2_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            dataGridView2.FirstDisplayedScrollingRowIndex = dataGridView2.Rows.Count - 1;
        }
        private void dataGridView3_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            dataGridView3.FirstDisplayedScrollingRowIndex = dataGridView3.Rows.Count - 1;
        }
        private void dataGridView4_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            dataGridView4.FirstDisplayedScrollingRowIndex = dataGridView4.Rows.Count - 1;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            OPC_disconnect();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            dataGridView2.ClearRows();
            dataGridView3.ClearRows();
            dataGridView4.ClearRows();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string value = SynReadOpcItem(textBox1.Text.Trim());
            MessageBox.Show(value);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bool temp = SynWriteOpcItem(textBox1.Text.Trim(), textBox2.Text.Trim());
            if (temp)
            {
                MessageBox.Show("写成功！");
            }
            else
            {
                MessageBox.Show("写失败！");
            }
        }
        #endregion

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
                        case "设备报警":
                            DeviceAlarm(address);
                            break;
                        case "设备状态":
                            DeviceState(address);
                            break;
                        case "交互地址":
                            DeviceInter(address);
                            break; //采集机床所有刀具信息


                    }
                }
            }
        }

        #region OPC相关函数方法
        /// <summary>
        /// OPC连接
        /// </summary>
        /// <param name="barcode"></param>
        private void OPC_connect()
        {

            #region 连接OPCSever
            //if (!OpcGroupClassm_server.IsConnected) { return; }
            //获取本地的IP地址
            var addr = KepIP;
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
        private void OPC_disconnect()
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
        private Item FindOpcItem(string itemName)
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
        private List<Item> FindOpcItems(List<string> itemNames)
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
        private bool SynWriteOpcItem(string itemName, string writeValue)
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
        private bool SynWriteOpcItems(Dictionary<string, string> writeItems)
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
        private string SynReadOpcItem(string itemName)
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
        private Dictionary<string, string> SynReadOpcItems(List<string> itemNames)
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

        private List<string> SynReadOpcItems2(List<string> itemNames)//同步读多个数据生成动态数组
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

        #region 数据采集方法
        private void DeviceAlarm(Test test)
        {
            this.dataGridView2.AddRow(test.Id, test.Value, test.Description);
        }
        private void DeviceState(Test test)
        {
            this.dataGridView3.AddRow(test.Id, test.Value, test.Description);
        }
        private void DeviceInter(Test test)
        {
            if (test.Content == "读成功")
            {
                if (test.Value == "1" || test.Value == "True")
                {
                    var BirthCodeToEquip = dicTest.FirstOrDefault(s => s.Value.Content == "读电子标签");
                    string vin = SynReadOpcItem(BirthCodeToEquip.Value.Id);
                    label1.SetText(vin);
                }
               
            }

            if (test.Content == "拧紧轴控制器工作合格")
            {
                if (test.Value == "1" || test.Value == "True")
                {
                    var BirthCodeToEquip = dicTest.FirstOrDefault(s => s.Value.Content == "拧紧轴角度值");
                    string angle = SynReadOpcItem(BirthCodeToEquip.Value.Id);
                    label6.SetText(angle);
                }
               
            }


            this.dataGridView4.AddRow(test.Id, test.Value, test.Description);
        }
        #endregion

        /// <summary>
        /// 获取本机的IP地址(IPv4地址)
        /// </summary>
        /// <returns></returns>
        public static string GetLocalIP()
        {
            try
            {
                IPHostEntry IpEntry = Dns.GetHostEntry(Dns.GetHostName());
                for (int i = 0; i < IpEntry.AddressList.Length; i++)
                {
                    if (IpEntry.AddressList[i].AddressFamily == AddressFamily.InterNetwork)
                    {
                        return IpEntry.AddressList[i].ToString();
                    }
                }
                return "";
            }
            catch (Exception ex)
            {
                MessageBox.Show("获取本机IP出错:" + ex.Message);
                return "";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="xmlPath"></param>
        public static string getKepIP(string xmlPath)
        {
            try
            {
                string IP = "";
                XmlDocument xml_doc = new XmlDocument();
                XmlReaderSettings settings = new XmlReaderSettings();
                settings.IgnoreComments = true;//忽略文档里面的注释
                XmlReader reader = XmlReader.Create(xmlPath, settings);
                xml_doc.Load(reader);
                XmlNode root_node = xml_doc.SelectSingleNode("data");
                XmlNodeList list_nodes = root_node.ChildNodes;
                foreach (XmlNode _nodes in list_nodes)
                {
                    switch (_nodes.Name)
                    {
                        case "KepWare":
                            IP = _nodes.InnerText;
                            break;
                    }
                }
                reader.Close();
                return IP;
            }
            catch (Exception ex) { throw; }
        }


    }
}
