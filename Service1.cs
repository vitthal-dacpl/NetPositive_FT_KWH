using Opc.Ua;
using Opc.Ua.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace NetPositive_FT_KWH
{
    public partial class Service1 : ServiceBase
    {
        #region Initialoization
        public static Session session;
        public static SqlConnection connection1 = null;
        public static SqlConnection connection2 = null;
        string connectionString1 = $"Data Source= ;Initial Catalog=Honeywell;Integrated Security=True";
        string connectionString2 = $"Data Source= ;Initial Catalog=Honeywell;Integrated Security=True";
        private static Timer timer2 = null;
        private static BusinessLayer bl = (BusinessLayer)null;
        public Dictionary<string, string> tagValues = null;

        #endregion
        public Service1()
        {
            connection1 = new SqlConnection(connectionString1);
            connection2 = new SqlConnection(connectionString2);
            bl = new BusinessLayer();
            InitializeComponent();
        }

        protected async override void OnStart(string[] args)
        {
            ManagementObjectSearcher MOS = new ManagementObjectSearcher("Select * From Win32_BaseBoard");
            string motherbord = "";
            foreach (ManagementObject getserial in MOS.Get())
            {
                motherbord = getserial["SerialNumber"].ToString();
            }
            if (motherbord == ".4F287T3.CNCMS0027J00EA.")
            {
                await CHECKOPCSERVER();
                ConnectOPCSession();

                if (session != null)
                    if (session.Connected)
                    {
                        await StartServiceAsync();
                        await ReadTagValuesAndSave();
                        timer2.Interval = 60000;
                        timer2.Elapsed += Timer2_Elapsed;
                        timer2.Enabled = true;
                    }
            }
            else
            {
                WriteToFile("You Have No Liecence Please Contact Admin");
            }
        }
        public async Task StartServiceAsync()
        {
            // Calculate the time until the next 7:00 AM
            DateTime now = DateTime.Now;
            DateTime targetTime = DateTime.Today.AddHours(7);

            // If it's already past 7:00 AM today, set target time to 7:00 AM the next day
            if (now > targetTime)
            {
                targetTime = targetTime.AddDays(1);
            }

            // Calculate the delay (time to wait until 7:00 AM)
            TimeSpan timeToWait = targetTime - now;


            // Wait until the next 7:00 AM
            await Task.Delay(timeToWait);

            // Read the values at 7:00 AM
            ReadValuesFromServer();
            //   await ReadTagValuesAndSave();
            // Set up the next reading at 7:00 AM the following day
            while (true)
            {
                // Wait another 24 hours (86400000 ms) for the next 7:00 AM
                await Task.Delay(TimeSpan.FromDays(1));

                // Read the values at 7:00 AM
                ReadValuesFromServer();
                // await  ReadTagValuesAndSave();
            }
        }
        static void MyHandler(object sender, UnhandledExceptionEventArgs args)
        {
            Exception e = (Exception)args.ExceptionObject;
            WriteToFile("Simple Service Error on: {0} " + e.Message + e.StackTrace);
        }
        public static bool PingHostA()
        {
            bool pingable = false;
            Ping pinger = null;
            try
            {
                pinger = new Ping();
                PingReply reply = pinger.Send("192.168.1.20");
                pingable = reply.Status == IPStatus.Success;
            }
            catch (PingException ex)
            {
                WriteToFile(ex.ToString() + "ping error" + DateTime.Now.ToString());
            }
            finally
            {
                if (pinger != null)
                {
                    pinger.Dispose();
                }
            }

            return pingable;
        }
        public static bool PingHostB()
        {
            bool pingable = false;
            Ping pinger = null;
            try
            {
                pinger = new Ping();
                PingReply reply = pinger.Send("192.168.1.21");
                pingable = reply.Status == IPStatus.Success;
            }
            catch (PingException ex)
            {
                WriteToFile(ex.ToString() + "ping error" + DateTime.Now.ToString());
            }
            finally
            {
                if (pinger != null)
                {
                    pinger.Dispose();
                }
            }

            return pingable;
        }
        private async void Timer2_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                NodeId Check_Connection = new NodeId("");
                DataValue dataValue = session.ReadValue(Check_Connection);
            }
            catch (Exception ex)
            {
                WriteToFile(ex.Message + " Timer2_Elapsed Reconnect " + DateTime.Now.ToString());
                await CHECKOPCSERVER();
            }
        }
        private static void WriteToFile(string text)
        {
            try
            {
                string path = "D:\\Honeywell_reports\\Report_services\\Report.txt";
                using (StreamWriter writer = new StreamWriter(path, true))
                {
                    writer.WriteLine(string.Format(text, DateTime.Now.ToString()));
                    writer.Close();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static async Task CHECKOPCSERVER()
        {
            try
            {
                if (session != null)
                {
                    session.Close();
                    session.Dispose();
                }
            }
            catch (Exception)
            {

                throw;
            }


            //////// OPC SERVER A/////////////

            try
            {

                if (PingHostA() == true)
                {
                    var endpointUrl = "opc.tcp://192.168.1.20:4840"; // Replace with your server's endpoint URL
                                                                     // Create a new application configuration
                    Utils.SetTraceOutput(Utils.TraceOutput.Off);
                    var config = new ApplicationConfiguration()
                    {
                        ServerConfiguration = new ServerConfiguration
                        {
                            UserTokenPolicies = new UserTokenPolicyCollection(new[] { new UserTokenPolicy(UserTokenType.UserName) }),
                        },
                        ApplicationName = "MyHomework",
                        ApplicationType = ApplicationType.Client,
                        SecurityConfiguration = new SecurityConfiguration
                        {
                            ApplicationCertificate = new CertificateIdentifier
                            {
                                StoreType = @"Windows",
                                StorePath = @"CurrentUser\My",
                                SubjectName = Utils.Format(@"CN={0}, DC={1}", "MyHomework", System.Net.Dns.GetHostName())
                            },
                            TrustedPeerCertificates = new CertificateTrustList
                            {
                                StoreType = @"Windows",
                                StorePath = @"CurrentUser\TrustedPeople",
                            },
                            NonceLength = 32,
                            AutoAcceptUntrustedCertificates = true
                        },
                        TransportConfigurations = new TransportConfigurationCollection(),
                        TransportQuotas = new TransportQuotas { OperationTimeout = 15000 },
                        ClientConfiguration = new ClientConfiguration { DefaultSessionTimeout = 60000, }
                    };

                    config.CertificateValidator = new CertificateValidator();
                    config.CertificateValidator.CertificateValidation += (s, certificateValidationEventArgs) =>
                    {
                        certificateValidationEventArgs.Accept = true; // Accept all certificates for testing purposes; modify this for production.
                    };

                    // Create a new session with the OPC UA server asynchronously
                    session = await Session.Create(config, new ConfiguredEndpoint(null, new EndpointDescription(endpointUrl)), true, "", 60000 * 60 * 24, new UserIdentity(), null);


                    if (session.Connected)
                    {
                        WriteToFile("OPC SERVER A Connected");

                    }
                }
                else
                {
                    WriteToFile("OPC SERVER A Not ping");
                    if (PingHostB() == true)
                    {

                        try
                        {
                            var endpointUrl = "opc.tcp://192.168.1.21:4840"; // Replace with your server's endpoint URL
                                                                             // Create a new application configuration
                            Utils.SetTraceOutput(Utils.TraceOutput.Off);
                            var config = new ApplicationConfiguration()
                            {
                                ServerConfiguration = new ServerConfiguration
                                {
                                    UserTokenPolicies = new UserTokenPolicyCollection(new[] { new UserTokenPolicy(UserTokenType.UserName) }),
                                },
                                ApplicationName = "MyHomework",
                                ApplicationType = ApplicationType.Client,
                                SecurityConfiguration = new SecurityConfiguration
                                {
                                    ApplicationCertificate = new CertificateIdentifier
                                    {
                                        StoreType = @"Windows",
                                        StorePath = @"CurrentUser\My",
                                        SubjectName = Utils.Format(@"CN={0}, DC={1}", "MyHomework", System.Net.Dns.GetHostName())
                                    },
                                    TrustedPeerCertificates = new CertificateTrustList
                                    {
                                        StoreType = @"Windows",
                                        StorePath = @"CurrentUser\TrustedPeople",
                                    },
                                    NonceLength = 32,
                                    AutoAcceptUntrustedCertificates = true
                                },
                                TransportConfigurations = new TransportConfigurationCollection(),
                                TransportQuotas = new TransportQuotas { OperationTimeout = 15000 },
                                ClientConfiguration = new ClientConfiguration { DefaultSessionTimeout = 60000, }
                            };

                            config.CertificateValidator = new CertificateValidator();
                            config.CertificateValidator.CertificateValidation += (s, certificateValidationEventArgs) =>
                            {
                                certificateValidationEventArgs.Accept = true; // Accept all certificates for testing purposes; modify this for production.
                            };

                            // Create a new session with the OPC UA server asynchronously
                            session = await Session.Create(config, new ConfiguredEndpoint(null, new EndpointDescription(endpointUrl)), true, "", 60000 * 60 * 24, new UserIdentity(), null);

                            if (session.Connected)
                            {
                                WriteToFile("OPC SERVER B Connected");
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteToFile(ex.Message + "  OPC SERVER B  Connect Error");
                        }
                    }
                    else
                    {
                        WriteToFile("OPC SERVER B Not Ping");
                    }

                }
            }
            catch (Exception ex)
            {

                WriteToFile(ex.Message + "  OPC SERVER A Connect Error");
                if (PingHostB() == true)
                {
                    try
                    {
                        var endpointUrl = "opc.tcp://192.168.1.21:4840"; // Replace with your server's endpoint URL
                                                                         // Create a new application configuration
                        Utils.SetTraceOutput(Utils.TraceOutput.Off);
                        var config = new ApplicationConfiguration()
                        {
                            ServerConfiguration = new ServerConfiguration
                            {
                                UserTokenPolicies = new UserTokenPolicyCollection(new[] { new UserTokenPolicy(UserTokenType.UserName) }),
                            },
                            ApplicationName = "MyHomework",
                            ApplicationType = ApplicationType.Client,
                            SecurityConfiguration = new SecurityConfiguration
                            {
                                ApplicationCertificate = new CertificateIdentifier
                                {
                                    StoreType = @"Windows",
                                    StorePath = @"CurrentUser\My",
                                    SubjectName = Utils.Format(@"CN={0}, DC={1}", "MyHomework", System.Net.Dns.GetHostName())
                                },
                                TrustedPeerCertificates = new CertificateTrustList
                                {
                                    StoreType = @"Windows",
                                    StorePath = @"CurrentUser\TrustedPeople",
                                },
                                NonceLength = 32,
                                AutoAcceptUntrustedCertificates = true
                            },
                            TransportConfigurations = new TransportConfigurationCollection(),
                            TransportQuotas = new TransportQuotas { OperationTimeout = 15000 },
                            ClientConfiguration = new ClientConfiguration { DefaultSessionTimeout = 60000, }
                        };

                        config.CertificateValidator = new CertificateValidator();
                        config.CertificateValidator.CertificateValidation += (s, certificateValidationEventArgs) =>
                        {
                            certificateValidationEventArgs.Accept = true; // Accept all certificates for testing purposes; modify this for production.
                        };

                        // Create a new session with the OPC UA server asynchronously
                        session = await Session.Create(config, new ConfiguredEndpoint(null, new EndpointDescription(endpointUrl)), true, "", 60000 * 60 * 24, new UserIdentity(), null);

                        if (session.Connected)
                        {
                            WriteToFile("OPC SERVER B Connected");
                        }
                    }
                    catch (Exception exx)
                    {
                        WriteToFile(exx.Message + "  OPC SERVER B  Connect Error");
                    }
                }
                else
                {
                    WriteToFile("OPC SERVER B Not Ping");
                }
            }
            ConnectOPCSession();
        }
        public static void ConnectOPCSession()
        {
            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.UnhandledException += new UnhandledExceptionEventHandler(MyHandler);

            try
            {

                if (session != null)
                {
                    if (session.Connected)
                    {
                        Subscription subscription = new Subscription(session.DefaultSubscription)
                        {
                            PublishingInterval = 1000, // in milliseconds
                            PublishingEnabled = true,
                            LifetimeCount = 100,
                            //   MaxKeepAliveCount = 10,
                            MaxNotificationsPerPublish = 100,
                            Priority = 1
                        };
                        session.AddSubscription(subscription);
                        subscription.Create();

                        // 
                        //MonitoredItem Reactor1 = new MonitoredItem(subscription.DefaultItem)
                        //{

                        //    DisplayName = "reportST_svp_pe_001",
                        //    StartNodeId = "ns=1;s=smvews:svp_pe_001_operation_fp.st_sp_report.pvfl", // Replace with the NodeId you want to monitor
                        //    AttributeId = Attributes.Value,
                        //    MonitoringMode = MonitoringMode.Reporting,
                        //    SamplingInterval = 1000 // in milliseconds,
                        //                            //LastValue = Opc.Ua.MonitoredItemNotification.
                        //};
                        //Reactor1.Notification += OnReactorNotification1;
                        //subscription.AddItem(Reactor1);

                        subscription.ApplyChanges();

                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile(ex.Message + " Monitered Item Error " + DateTime.Now.ToString());
            }

        }
        public static object ReadNodes(NodeId nodeID)
        {
            try
            {
                DataValue value = session.ReadValue(nodeID);
                object obj = value.Value;
                return obj;
            }
            catch (Exception ex)
            {
                WriteToFile(ex.Message + " ReadNode " + " " + nodeID + " " + DateTime.Now.ToString());

                return null;
            }

        }
        public static void ReadValuesFromServer()
        {
            NodeId NodeEm1 = new NodeId("");
            NodeId NodeEm2 = new NodeId("");
            NodeId NodeEm3 = new NodeId("");
            NodeId NodeEm4 = new NodeId("");
            NodeId NodeEm5 = new NodeId("");
            NodeId NodeEm6 = new NodeId("");
            NodeId NodeEm7 = new NodeId("");
            NodeId NodeEm8 = new NodeId("");
            NodeId NodeEm9 = new NodeId("");
            NodeId NodeEm10 = new NodeId("");
            NodeId NodeEm11 = new NodeId("");
            NodeId NodeEm12 = new NodeId("");
            NodeId NodeEm13 = new NodeId("");
            NodeId NodeEm14 = new NodeId("");
            NodeId NodeEm15 = new NodeId("");
            NodeId NodeEm16 = new NodeId("");
            NodeId NodeEm17 = new NodeId("");
            NodeId NodeEm18 = new NodeId("");
            NodeId NodeEm19 = new NodeId("");
            NodeId NodeEm20 = new NodeId("");
            NodeId NodeEm21 = new NodeId("");
            NodeId NodeEm22 = new NodeId("");
            NodeId NodeEm23 = new NodeId("");
            NodeId NodeEm24 = new NodeId("");
            NodeId NodeEm25 = new NodeId("");
            NodeId NodeEm26 = new NodeId("");
            NodeId NodeEm27 = new NodeId("");
            NodeId NodeEm28 = new NodeId("");
            NodeId NodeEm29 = new NodeId("");
            NodeId NodeEm30 = new NodeId("");
            NodeId NodeFm1 = new NodeId("");
            NodeId NodeFm2 = new NodeId("");
            NodeId NodeFm3 = new NodeId("");
            NodeId NodeFm4 = new NodeId("");
            NodeId NodeFm5 = new NodeId("");
            NodeId NodeFm6 = new NodeId("");
            NodeId NodeFm7 = new NodeId("");
            NodeId NodeFm8 = new NodeId("");
            NodeId NodeFm9 = new NodeId("");
            NodeId NodeFm10 = new NodeId("");
            NodeId NodeFm11 = new NodeId("");
            NodeId NodeFm12 = new NodeId("");
            NodeId NodeFm13 = new NodeId("");
            NodeId NodeFm14 = new NodeId("");
            NodeId NodeFm15 = new NodeId("");
            NodeId NodeFm16 = new NodeId("");
            NodeId NodeFm17 = new NodeId("");
            NodeId NodeFm18 = new NodeId("");
            NodeId NodeFm19 = new NodeId("");
            NodeId NodeFm20 = new NodeId("");
            object Em1 = ReadNodes(NodeEm1);
            object Em2 = ReadNodes(NodeEm2);
            object Em3 = ReadNodes(NodeEm3);
            object Em4 = ReadNodes(NodeEm4);
            object Em5 = ReadNodes(NodeEm5);
            object Em6 = ReadNodes(NodeEm6);
            object Em7 = ReadNodes(NodeEm7);
            object Em8 = ReadNodes(NodeEm8);
            object Em9 = ReadNodes(NodeEm9);
            object Em10 = ReadNodes(NodeEm10);
            object Em11 = ReadNodes(NodeEm11);
            object Em12 = ReadNodes(NodeEm12);
            object Em13 = ReadNodes(NodeEm13);
            object Em14 = ReadNodes(NodeEm14);
            object Em15 = ReadNodes(NodeEm15);
            object Em16 = ReadNodes(NodeEm16);
            object Em17 = ReadNodes(NodeEm17);
            object Em18 = ReadNodes(NodeEm18);
            object Em19 = ReadNodes(NodeEm19);
            object Em20 = ReadNodes(NodeEm20);
            object Em21 = ReadNodes(NodeEm21);
            object Em22 = ReadNodes(NodeEm22);
            object Em23 = ReadNodes(NodeEm23);
            object Em24 = ReadNodes(NodeEm24);
            object Em25 = ReadNodes(NodeEm25);
            object Em26 = ReadNodes(NodeEm26);
            object Em27 = ReadNodes(NodeEm27);
            object Em28 = ReadNodes(NodeEm28);
            object Em29 = ReadNodes(NodeEm29);
            object Em30 = ReadNodes(NodeEm30);
            object Fm1 = ReadNodes(NodeFm1);
            object Fm2 = ReadNodes(NodeFm2);
            object Fm3 = ReadNodes(NodeFm3);
            object Fm4 = ReadNodes(NodeFm4);
            object Fm5 = ReadNodes(NodeFm5);
            object Fm6 = ReadNodes(NodeFm6);
            object Fm7 = ReadNodes(NodeFm7);
            object Fm8 = ReadNodes(NodeFm8);
            object Fm9 = ReadNodes(NodeFm9);
            object Fm10 = ReadNodes(NodeFm10);
            object Fm11 = ReadNodes(NodeFm11);
            object Fm12 = ReadNodes(NodeFm12);
            object Fm13 = ReadNodes(NodeFm13);
            object Fm14 = ReadNodes(NodeFm14);
            object Fm15 = ReadNodes(NodeFm15);
            object Fm16 = ReadNodes(NodeFm16);
            object Fm17 = ReadNodes(NodeFm17);
            object Fm18 = ReadNodes(NodeFm18);
            object Fm19 = ReadNodes(NodeFm19);
            object Fm20 = ReadNodes(NodeFm20);

            bl.InsertData(
                (object)Em1, (object)Em2, (object)Em3, (object)Em4, (object)Em5, (object)Em6, (object)Em7, (object)Em8,
                (object)Em9, (object)Em10, (object)Em11, (object)Em12, (object)Em13, (object)Em14, (object)Em15,
                (object)Em16, (object)Em17, (object)Em18, (object)Em19, (object)Em20, (object)Em21, (object)Em22,
                (object)Em23, (object)Em24, (object)Em25, (object)Em26, (object)Em27, (object)Em28, (object)Em29,
                (object)Em30, (object)Fm1, (object)Fm2, (object)Fm3, (object)Fm4, (object)Fm5, (object)Fm6,
                (object)Fm7, (object)Fm8, (object)Fm9, (object)Fm10, (object)Fm11, (object)Fm12, (object)Fm13,
                (object)Fm14, (object)Fm15, (object)Fm16, (object)Fm17, (object)Fm18, (object)Fm19, (object)Fm20

            );

        }


        // Method to fetch tags from the database dynamically


        // Method to read tag values dynamically
        public async Task ReadTagValuesAndSave()
        {
            // Fetch tags dynamically from database (or fixed list if you prefer)
            List<string> tags = bl.Gettags(); // Assume you fetch tag names from DB or predefined list

            // Dictionary to store tag names and their corresponding values
            tagValues = new Dictionary<string, string>();

            // Read each tag's value from the OPC UA server using a for loop
            for (int i = 0; i < tags.Count; i++)
            {
                try
                {
                    // Read the tag value from the OPC UA server
                    var value = await session.ReadValueAsync(tags[i]); // tags[i] is the current tag

                    // Store the tag value in the dictionary
                    tagValues[$"@EM{i + 1}"] = value.ToString();
                }
                catch (Exception)
                {
                    tagValues[$"@EM{i + 1}"] = " "; // Storing error for failed tags
                }
            }

            // Store the tag values in SQL Server in a single row
            SaveTagValuesToDatabase(tagValues);
        }


        // Method to save tag values into SQL Server in a single row
        private void SaveTagValuesToDatabase(Dictionary<string, string> tagvalues)
        {
            int a = bl.InsertData(tagvalues);
            #region Direct insert
            //string connectionString = ConfigurationManager.AppSettings["Sqlconnstring"].ToString();
            // using (SqlConnection conn = new SqlConnection(connectionString))
            // {
            //     conn.Open();

            //     // Prepare the SQL Insert query with 50 parameters
            //     string query = @"
            //     INSERT INTO TagValues (
            //         Em1, Em2, Em3, Em4, Em5, Em6, Em7, Em8, Em9, Em10, Em11, Em12, Em13, Em14, Em15, Em16, Em17, Em18, Em19, Em20,
            //         Em21, Em22, Em23, Em24, Em25, Em26, Em27, Em28, Em29, Em30, Em31, Em32, Em33, Em34, Em35, Em36, Em37, Em38, Em39, Em40,
            //         Em41, Em42, Em43, Em44, Em45, Em46, Em47, Em48, Em49, Em50
            //     ) 
            //     VALUES (
            //         @Em1, @Em2, @Em3, @Em4, @Em5, @Em6, @Em7, @Em8, @Em9, @Em10, @Em11, @Em12, @Em13, @Em14, @Em15, @Em16, @Em17, @Em18, @Em19, @Em20,
            //         @Em21, @Em22, @Em23, @Em24, @Em25, @Em26, @Em27, @Em28, @Em29, @Em30, @Em31, @Em32, @Em33, @Em34, @Em35, @Em36, @Em37, @Em38, @Em39, @Em40,
            //         @Em41, @Em42, @Em43, @Em44, @Em45, @Em46, @Em47, @Em48, @Em49, @Em50
            //     )";

            //     using (SqlCommand cmd = new SqlCommand(query, conn))
            //     {
            //         // Add parameters dynamically for each tag
            //         for (int i = 1; i <= 50; i++)
            //         {
            //             string tagName = $"Em{i}";
            //             cmd.Parameters.AddWithValue($"@{tagName}", tagValues.ContainsKey(tagName) ? (object)tagValues[tagName] : DBNull.Value);
            //         }



            //         // Execute the insert query
            //         cmd.ExecuteNonQuery();
            //     }
            // }
            #endregion
            if (a != 0)
            {
                tagValues.Clear();
            }
            else
            {
                SaveTagValuesToDatabase(tagValues);
            }
        }

        protected override void OnStop()
        {
            try
            {
                if (session != null)
                {
                    session.Close();
                    session.Dispose();
                }
                timer2 = null;
                connection1 = null;
                connection2 = null;
                WriteToFile("Service Stop");
            }
            catch (Exception ex)
            {
                WriteToFile("Service Stop at Exception" + ex.ToString());
            }
        }
    }
}
