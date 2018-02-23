using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DBUtility;
using System.Configuration;
using WeYuDB;
using WeYuFunctionLibrary;
using log4net;
using System.IO;
using System.Diagnostics;
using System.Collections;
using System.Reflection;

namespace ZZ_WoScheduleLoader
{
    public partial class Monitor : Form
    {

        public string assemblyVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
        public ILog log;

        //App.config當中的參數在此作封裝。
        public string Log4ConfigFilePath{get{return AppDomain.CurrentDomain.BaseDirectory + @"\\" + ConfigurationManager.AppSettings["Log4ConfigFileName"];}}
        public int DATA_DAY_RANGE{get{return ConfigurationManager.AppSettings["DATA_DAY_RANGE"].ToInt16OrDefault(2);}}

        public string OracleWOTableName{get{return ConfigurationManager.AppSettings["OracleWOTableName"].ToString();}}
        public string OracleAcceptOrderTableName{get{return ConfigurationManager.AppSettings["OracleAcceptOrderTableName"].ToString();}}
        public string OracleScheduleTableName{get{return ConfigurationManager.AppSettings["OracleScheduleTableName"].ToString();}}
        public string OracleScheduleTableFields{get{return ConfigurationManager.AppSettings["OracleScheduleTableFields"].ToString();}}
        public string OracleScheduleTableSortField{get{return ConfigurationManager.AppSettings["OracleScheduleTableSortField"].ToString();}}       

        public string SQLServerWOTableName{get{return ConfigurationManager.AppSettings["SQLServerWOTableName"].ToString();}}
        public string SQLServerScheduleTableName{get{return ConfigurationManager.AppSettings["SQLServerScheduleTableName"].ToString();}}
        public string SQLServerScheduleTableFields{get{return ConfigurationManager.AppSettings["SQLServerScheduleTableFields"].ToString();}}
        public string SQLServerScheduleTableSortField {get{return ConfigurationManager.AppSettings["SQLServerScheduleTableSortField"].ToString();}}
        

        public string ScheduleTableKeyFields_Date {get{return ConfigurationManager.AppSettings["ScheduleTableKeyFields_Date"].ToString();}}
        public string ScheduleTableKeyFields_String {get{return ConfigurationManager.AppSettings["ScheduleTableKeyFields_String"].ToString();}}
        public string ScheduleTableKeyFields_Number {get{return ConfigurationManager.AppSettings["ScheduleTableKeyFields_Number"].ToString();}}


        public Monitor()
        {

            log4net.Config.XmlConfigurator.ConfigureAndWatch(new System.IO.FileInfo(Log4ConfigFilePath));
            log = LogManager.GetLogger(typeof(Monitor));

            //InitializeComponent();   //顯示UI介面(要以手動方式執行button1時才可能需要顯示)
            button1_Click(this, null); //按鈕1的執行。

        }
        
        private void Monitor_Load(object sender, EventArgs e)
        {

        }
        
        private void button1_Click(object sender, EventArgs e)
        {

            var Totalclock = new Stopwatch();
            Totalclock.Start();  //---總時鐘開始----
            
            try
            {
            
                log.Info("版本號:" + assemblyVersion);

                CreateWO(); //負責「創建批號」(其實應該叫創建工單WO)的Function

                CreateSchedule(); //負責創建排程資料的Function
            
            }
            catch (Exception e1)
            {
                log.Info(e1.Message);
                log.Info("******Loader於執行時發生錯誤，執行終止!!*******");
            }
            finally
            {
                Totalclock.Stop();
                log.Info("總共耗時為:" + Totalclock.Elapsed.TotalSeconds.ToString() + "秒");
                log.Info("");

                System.Environment.Exit(0); //最後要關閉整個黑畫面。
            }
                       
        }

        //負責「創建批號」(其實應該叫創建工單WO)的Function
        private void CreateWO() {

            log.Info("自動建立工單程式，執行開始!");
            log.Info("OracleWOTableName:" + OracleWOTableName + ",SQLServerWOTableName:" + SQLServerWOTableName + ",DATA_DAY_RANGE:" + DATA_DAY_RANGE);

            OracleDBUtility odb = new OracleDBUtility();
            DBUtility.DBUtility sdb = new DBUtility.DBUtility();

            var clock = new Stopwatch();
            //---計時鐘開始----
            clock.Start();

            //找出所有時間條件內的工單
            string sql = @"SELECT * FROM " + OracleWOTableName + " WHERE ROUND(TO_NUMBER(sysdate - UPDATE_TIME))<=" + DATA_DAY_RANGE.ToString() + " and (TXT04='REL' OR TXT04='LKD' ) order by UPDATE_TIME desc";

            log.Info("主要SQL指令: " + sql);  //←本程式最關鍵的SQL指令(印出來看-可直接複製來作驗證使用)。


            int count = 0; //初始化計數器

            DataTable dtOracleWOs = odb.OracleGetDataTable(sql);
            for (int i = 0; i < dtOracleWOs.Rows.Count; i++)
            {
                //System.Threading.Thread.Sleep(5);

                //把Oracle資料庫當中的每筆工單，逐一取出
                string WO = dtOracleWOs.Rows[i]["PRODUCE_NO"].ToString();

                //再拿WO去SQL Server資料庫裡免檢驗，是否已經存在
                sql = "select WO from " + SQLServerWOTableName + " where WO ='" + WO + "' ";
                DataTable dtSQLData = sdb.GetReader(sql);

                //如果查出來的結果，WO還沒有存在SQL Server裡面，就CreateLot(建立批號)
                if (dtSQLData == null || dtSQLData.Rows.Count == 0)
                {
                    log.Info("=====工單: " + WO + " 還不存在於 " + SQLServerWOTableName + " 當中=====");  //寫入Log

                    //在就CreateLot之前，要執行-協助寫入WOR_MASTER及工序相關的function(SyncWoRoute)。
                    SyncWoRoute(dtOracleWOs.Rows[i]);

                    #region ====建立批號===============
                    //log.Info(WO + " 開始「創建批號」!!");

                    try
                    {
                        bool IsMutiRoute = false;
                        #region 判断是一般流程还是平行流程
                        sql = @"select MIN(VORNR1) AS VORNR1,MIN(VORNR2) AS VORNR2,APLZL,PLNFL,FLGAT 
                                    from ZZ_CYM_PP_ORDER_SEQUE 
                                    where produce_no='" + WO + @"'
                                    GROUP BY APLZL,PLNFL,FLGAT";
                        DataTable dtSEQU = sdb.GetReader(sql);
                        if (dtSEQU != null && dtSEQU.Rows != null && dtSEQU.Rows.Count > 1)
                        {
                            IsMutiRoute = true;
                        }
                        #endregion

                        if (IsMutiRoute)
                        {
                            log.Info(WO + " 平行流程開始產生批號");
                                DBQuery query = new DBQuery();
                                WipFunction wipfun = new WipFunction(query);
                                DataTable dtWO = wipfun.GetWoInfo(WO);
                                DataTable dtLotStatus = wipfun.GetLotStatus("Wait");
                                for (int x = 0; x < dtSEQU.Rows.Count; x++)
                                {
                                    bool bCreate = false;
                                    string RouteCode = "";
                                    string LotNo = WO;

                                    #region 判断平行流程是否需要产生批号
                                    if (dtSEQU.Rows[x]["VORNR1"].ToString().Trim().Equals(""))
                                    {
                                        bCreate = true;
                                        RouteCode = WO + "-" + dtSEQU.Rows[x]["PLNFL"].ToString().Substring(4, 2);
                                    }
                                    else
                                    {
                                        if (dtSEQU.Rows[x]["VORNR1"].ToString().Equals("0010"))
                                        {
                                            bCreate = true;
                                            RouteCode = WO + "-" + dtSEQU.Rows[x]["PLNFL"].ToString().Substring(4, 2);

                                            string zeroTester = dtSEQU.Rows[x]["PLNFL"].ToString().Substring(3, 3);
                                            if (zeroTester.Equals("000")) {
                                                log.Info("途程[" + RouteCode + "] 對應所創建的批號→LotNo[" + LotNo + "]");
                                                //如果發現PLNFL的末三碼是000，就不用生成.000的批號了。
                                                continue;
                                            }

                                            LotNo = WO + "." + dtSEQU.Rows[x]["PLNFL"].ToString().Substring(3, 3);

                                            log.Info("途程[" + RouteCode + "] 對應所創建的批號→LotNo[" + LotNo + "]");
                                        }
                                    }
                                    #endregion

                                    if (!bCreate)
                                        continue;

                                    sql = @"SELECT TOP 1 * FROM V_BAS_ROUTE_OPER WHERE ROUTE_CODE='" + RouteCode + "' ORDER BY SEQ ";
                                    DataTable dtStartRouteOper = query.Get_Reader(sql);
                                    DataTable dtFact = wipfun.GetFactoryBySID(dtWO.Rows[0]["FACTORY_SID"].ToString());
                                    DataTable dtInputUser = query.GetTable("SEC_USER", "ACCOUNT_NO", "ADMINV2");
                                    decimal HIST_SID = query.GetDBSid();
                                    decimal DATA_LINK_SID = query.GetDBSid();
                                    decimal LOT_SID = query.GetDBSid();
                                    bool bResult = wipfun.CreateLot(HIST_SID, DATA_LINK_SID, LOT_SID, LotNo, "", ""
                                                       , dtWO, dtLotStatus, dtStartRouteOper, dtWO.Rows[0]["QUANTITY"].ToString(), System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), dtFact, dtInputUser,
                                                       Convert.ToDecimal(dtStartRouteOper.Rows[0]["ROUTE_SID"].ToString()), "Loader");
                                    if (bResult)
                                    {
                                        log.Info("工單[" + WO + "]產生批號成功!");
                                    }
                                    else
                                    {
                                        log.Info("工單[" + WO + "]產生批號失敗!");
                                    }
                                }                      
                        }
                        else
                        {
                            log.Info(WO + " 一般流程開始產生批號");
                            DBQuery query = new DBQuery();
                            WipFunction wipfun = new WipFunction(query);
                            DataTable dtWO = wipfun.GetWoInfo(WO);
                            DataTable dtLotStatus = wipfun.GetLotStatus("Wait");
                            sql = @"SELECT TOP 1 * FROM V_BAS_ROUTE_OPER WHERE ROUTE_CODE='" + WO + "' ORDER BY SEQ ";
                            DataTable dtStartRouteOper = query.Get_Reader(sql);
                            DataTable dtFact = wipfun.GetFactoryBySID(dtWO.Rows[0]["FACTORY_SID"].ToString());

                            if (dtStartRouteOper == null || dtStartRouteOper.Rows.Count == 0)
                            {
                                log.Info(WO + "於V_BAS_ROUTE_OPER的查詢結果為空!原因是它在PP_ACCEPT_ORDER中找不到任何ZORDUPDFG='C'的工序資料可以拋給MES。奕璋經裡表示，這樣就可以當作本工單已經取消、不須理會。因此程式判定自動跳過不進行創建!!!");
                                log.Info("========================================================");
                                continue;
                            }

                            DataTable dtInputUser = query.GetTable("SEC_USER", "ACCOUNT_NO", "ADMINV2");
                            decimal HIST_SID = query.GetDBSid();
                            decimal DATA_LINK_SID = query.GetDBSid();
                            decimal LOT_SID = query.GetDBSid();
                            bool bResult = wipfun.CreateLot(HIST_SID, DATA_LINK_SID, LOT_SID, WO, "", ""
                                               , dtWO, dtLotStatus, dtStartRouteOper, dtWO.Rows[0]["QUANTITY"].ToString(), System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), dtFact, dtInputUser,
                                               Convert.ToDecimal(dtStartRouteOper.Rows[0]["ROUTE_SID"].ToString()), "Loader");
                            if (bResult)
                            {
                                log.Info("工單[" + WO + "]產生批號成功!");
                            }
                            else
                            {
                                log.Info("工單[" + WO + "]產生批號失敗!");
                            }
                        }
                    }
                    catch (Exception ex2)
                    {
                        log.Error(ex2.Message , ex2);
                    }
                    #endregion

                }
                //反之(已經存在)就寫log訊息提醒
                else
                {
                    log.Info(WO + " 在 " + SQLServerWOTableName + " 當中已經有資料，因此不再執行批號創建。");  //寫入Log
                }
                count++; //計數器+1

            }

            clock.Stop(); //批號創建的計時停止
            log.Info("");
            log.Info("耗用時間為:" + clock.Elapsed.TotalSeconds.ToString() + "秒");
            log.Info("-----批號創建比對的工單筆數有:" + count + "筆。-----");

        }

        //協助寫入WOR_MASTER及工序相關的function
        private void SyncWoRoute(DataRow toSyncWODataRow)
        {
            string WO = toSyncWODataRow["PRODUCE_NO"].ToString();

            OracleDBUtility odb = new OracleDBUtility();
            DBUtility.DBUtility sdb = new DBUtility.DBUtility();
            ArrayList sqls = new ArrayList();

            string sql = "select * from WOR_MASTER WHERE WO='" + toSyncWODataRow["PRODUCE_NO"].ToString() + "'";
            DataTable dtCheck = sdb.GetReader(sql);

            //如果WOR_MASTER裡面還沒有這張工單，就執行Insert。
            if (dtCheck == null || dtCheck.Rows == null || dtCheck.Rows.Count == 0)
            {
                #region ====BAS_FACTORY驗證========

                string FactoryCode = toSyncWODataRow["WERKS"].ToString();
                string FactorySid = "";
                sql = "select * from BAS_FACTORY WHERE FACTORY_CODE='" + FactoryCode + "' ";
                DataTable dtFactory = sdb.GetReader(sql);

                //如果BAS_FACTORY不存在，就手動創建。(目前看來不太有機會走進這裡，目前為止BAS_FACTORY都是用固定那兩筆資料)
                if (dtFactory == null || dtFactory.Rows == null || dtFactory.Rows.Count == 0)
                {
                    FactorySid = sdb.GetSid();
                    sql = @"insert into BAS_FACTORY (FACTORY_SID,FACTORY_CODE,FACTORY_NAME,ENABLE_FLAG,CREATE_USER,CREATE_TIME,EDIT_USER,EDIT_TIME)
                            VALUES
                            (" + FactorySid + ",'" + FactoryCode + "','" + FactoryCode + "','Y','Loader',getdate(),'Loader',getdate())";
                    sdb.doExecute(sql);
                }
                //幾乎永遠用不到上面的條件，而會走到這裡取得FACTORY_SID=123972402940953的那筆資料來用。
                else
                {
                    FactorySid = dtFactory.Rows[0]["FACTORY_SID"].ToString();
                }
                #endregion

                #region ====BAS_PARTNO物料主檔驗證===========

                string PartNo = toSyncWODataRow["CARD_NO"].ToString();
                string PartSid = "";
                sql = "SELECT * FROM BAS_PARTNO where PART_NO='" + PartNo + "'";
                DataTable dtPartNO = sdb.GetReader(sql);

                //如果BAS_PARTNO物料主檔不存在，就手動創建。
                if (dtPartNO == null || dtPartNO.Rows == null || dtPartNO.Rows.Count == 0)
                {
                    PartSid = sdb.GetSid();
                    sql = @"insert into BAS_PARTNO (PART_SID,PART_NO,PART_SPEC,ENABLE_FLAG,CREATE_USER,CREATE_TIME,EDIT_USER,EDIT_TIME)
                            VALUES
                            (" + PartSid + ",'" + PartNo + "','" + PartNo + "','Y', 'Loader',getdate(),'Loader',getdate())";
                    sdb.doExecute(sql);
                }
                //已經有的話，就拿現有的繼續往下走。
                else
                {
                    PartSid = dtPartNO.Rows[0]["PART_SID"].ToString();
                }
                #endregion

                string SID = sdb.GetSid();

                #region =====寫入資料進 WOR_MASTER==============

                //WO = toSyncWODataRow["PRODUCE_NO"].ToString();
                try
                {
                    string tmpUnit = toSyncWODataRow["GMEIN"].ToString();
                    if (tmpUnit.Equals("ST"))
                    {
                        tmpUnit = "PC";
                    }

                    sql = @"insert into WOR_MASTER (
                                    WO_SID,
                                    WO,
                                    WO_TYPE,
                                    ERP_WO_TYPE,
                                    PART_SID,
                                    PART_NO,
                                    ERP_RELEASE_QTY,
                                    QUANTITY,
                                    CREATE_USER,
                                    CREATE_TIME,
                                    EDIT_USER,
                                    EDIT_TIME,
                                    ERP_STATUS,
                                    FACTORY_SID,
                                    ACTIVE_DATE,
                                    UNIT,
                                    ERP_SCHEDULE_RELEASE_DATE,
                                    ERP_SCHEDULE_Finsh_DATE,
                                    MASTER_WO) VALUES (
                                    " + SID + @",
                                    '" + toSyncWODataRow["PRODUCE_NO"].ToString() + @"',
                                    '" + toSyncWODataRow["CATEGORY"].ToString() + @"',
                                    '" + toSyncWODataRow["CATEGORY"].ToString() + @"',
                                    '" + PartSid + @"',
                                    '" + toSyncWODataRow["CARD_NO"].ToString() + @"',
                                    " + toSyncWODataRow["PRODUCE_QTY"].ToString() + @",
                                    " + toSyncWODataRow["PRODUCE_QTY"].ToString() + @",
                                    '" + toSyncWODataRow["INSERT_USER"].ToString() + @"',
                                    getdate(),
                                    '" + toSyncWODataRow["UPDATE_USER"].ToString() + @"',
                                    getdate(),
                                    '" + toSyncWODataRow["TXT04"].ToString() + @"',
                                    '" + FactorySid + @"',";
                    if (toSyncWODataRow["FTRMI"].GetType().Name == "DBNull")
                        sql += @"null,
                                 '" + tmpUnit + @"',";
                    else
                        sql += @" '" + Convert.ToDateTime(toSyncWODataRow["FTRMI"]).ToString("yyyy-MM-dd HH:mm:ss") + @"',
                                    '" + tmpUnit + @"',";

                    if (toSyncWODataRow["GSTRS"].GetType().Name == "DBNull")
                        sql += "null,";
                    else
                        sql += @"'" + Convert.ToDateTime(toSyncWODataRow["GSTRS"]).ToString("yyyy-MM-dd HH:mm:ss") + @"',";

                    if (toSyncWODataRow["GLTRS"].GetType().Name == "DBNull")
                        sql += @"null,
                                    '" + toSyncWODataRow["MAUFNR"].ToString() + @"'
                                    )";
                    else
                        sql += @"'" + Convert.ToDateTime(toSyncWODataRow["GLTRS"]).ToString("yyyy-MM-dd HH:mm:ss") + @"',
                                '" + toSyncWODataRow["MAUFNR"].ToString() + @"'
                                )";
                    sqls.Add(sql);
                }
                catch (Exception ewo)
                {
                    log.Info(ewo.Message);
                }
                #endregion

                #region ===== 將PP_ACCEPT_ORDER查出來的資料 轉為→ 工序主檔 BAS_OPERATION==================

                //查出PP_ACCEPT_ORDER內容的SQL
                sql = @"select * from " + OracleAcceptOrderTableName + " where aufnr='" + WO + "' AND ZORDUPDFG='C' "; //增加STEUS<>'PP02'条件.  PP02代表委外作业
                DataTable dtWOOpers = odb.OracleGetDataTable(sql);

                //如果PP_ACCEPT_ORDER的內容是有東西的，就繼續往下走
                if (dtWOOpers != null && dtWOOpers.Rows != null && dtWOOpers.Rows.Count > 0)
                {
                    #region ===== 寫資料進 途程主檔 BAS_ROUTE==================

                    string RouteSID = sdb.GetSid();
                    sql = @"INSERT INTO BAS_ROUTE
                                           (ROUTE_SID
                                           ,ROUTE_CODE
                                           ,ROUTE_NAME
                                           ,LINK_TYPE
                                           ,DESCRIPTION
                                           ,ENABLE_FLAG
                                           ,CREATE_USER
                                           ,CREATE_TIME
                                           ,EDIT_USER
                                           ,EDIT_TIME)
                                     VALUES
                                           (" + RouteSID + @"
                                           ,'" + WO + @"'
                                           ,'" + WO + @"'
                                           ,'WO'
                                            ,'Route for " + WO + @"'
                                           ,'Y'
                                           ,'" + toSyncWODataRow["INSERT_USER"].ToString() + @"'
                                           ,getdate()
                                           ,'" + toSyncWODataRow["UPDATE_USER"].ToString() + @"'
                                           ,getdate()
		                                   )";
                    sqls.Add(sql);
                    #endregion

                    int PP02OperSEQ = 1;
                    #region ====== 將PP_ACCEPT_ORDER查出來的資料，逐一比對~ ================
                    for (int j = 0; j < dtWOOpers.Rows.Count; j++)
                    {
                        string OperationCode = "";
                        string OperationName = "";
                        string SapCode = "";

                        //如果STEUS是PP02，代表委外作業。
                        if (dtWOOpers.Rows[j]["STEUS"].ToString().Equals("PP02"))
                        {
                            //OperationCode = "OUTSOURCING" + PP02OperSEQ.ToString();
                            OperationCode = WO + "-OUTSOURCING" + PP02OperSEQ.ToString();
                            OperationName = dtWOOpers.Rows[j]["LTXA1"].ToString().Replace("'","");
                            SapCode = "OUTSOURCING" + PP02OperSEQ.ToString();
                            PP02OperSEQ++;
                        }
                        else
                        {
                            //OperationCode = dtWOOpers.Rows[j]["ARBPL"].ToString();
                            OperationCode = WO + "-" + dtWOOpers.Rows[j]["VORNR"].ToString();
                            OperationName = dtWOOpers.Rows[j]["LTXA1"].ToString().Replace("'", "");
                            SapCode = dtWOOpers.Rows[j]["ARBPL"].ToString();
                        }
                        string OperSID = sdb.GetSid();

                        #region ===== 寫入 工序主檔 BAS_OPERATION==================

                        sql = @"INSERT INTO BAS_OPERATION
                                               (OPERATION_SID
                                               ,OPERATION_CODE
                                               ,OPERATION_NAME
                                               ,SAP_CODE
                                               ,THROW_SAP
                                               ,CONTROL_MODE
                                               ,DESCRIPTION
                                               ,ENABLE_FLAG
                                               ,CREATE_USER
                                               ,CREATE_TIME
                                               ,EDIT_USER
                                               ,EDIT_TIME)
                                         VALUES
                                               (" + OperSID + @"
                                               ,'" + OperationCode + @"'
                                               ,'" + OperationName + @"'
                                               ,'" + SapCode + @"'
                                               ,'Y'
                                               ,''
                                               ,'" + OperationName + @"'
                                               ,'Y'
                                               ,'" + toSyncWODataRow["INSERT_USER"].ToString() + @"'
                                               ,getdate()
                                               ,'" + toSyncWODataRow["UPDATE_USER"].ToString() + @"'
                                               ,getdate()
		                                       )
                                    ";
                        sqls.Add(sql);
                        #endregion

                        #region ===== 寫入 途程與工序對應 BAS_ROUTE_OPER==================

                        string tmpUnit = dtWOOpers.Rows[j]["GMEIN"].ToString();
                        if (tmpUnit.Equals("ST"))
                        {
                            tmpUnit = "PC";
                        }
                        sql = @"INSERT INTO BAS_ROUTE_OPER
                                           (ROUTE_OPER_SID
                                           ,ROUTE_SID
                                           ,SEQ
                                           ,OPERATION_SID
                                           ,SAP_SEQ
                                           ,SAP_CODE
                                           ,THROW_SAP
                                           ,UNIT
                                           ,CREATE_USER
                                           ,CREATE_TIME
                                           ,EDIT_USER
                                           ,EDIT_TIME
                                          )
                                     VALUES
                                           (dbo.GetSid()
                                           ," + RouteSID + @"
                                           ," + (j + 1).ToString() + @"
                                           ," + OperSID + @"
                                           ,'" + dtWOOpers.Rows[j]["VORNR"].ToString() + @"'
                                           ,'" + SapCode + @"'
                                           ,''
                                           ,'" + tmpUnit + @"'
                                           ,'" + toSyncWODataRow["INSERT_USER"].ToString() + @"'
                                           ,getdate()
                                           ,'" + toSyncWODataRow["UPDATE_USER"].ToString() + @"'
                                           ,getdate()
		                                   )
                                ";
                        sqls.Add(sql);
                        #endregion

                        System.Threading.Thread.Sleep(5);
                    }
                    #endregion
                }
                else
                {
                    log.Info("未找到工單[" + WO + "]的作業站記錄 (亦即PP_ACCEPT_ORDER where aufnr='" + WO + "' AND ZORDUPDFG='C' IS NULL)");
                }

                #endregion

            }
            else
            {
                log.Info("[" + WO + "]在WOR_MASTER中已經有資料，因此本Loader不再重複Insert(Update由廠商的Loader來處理)");
            }

            if (sdb.ExecuteTransation(sqls))
            {
                //log.Info("[" + WO + "]已完成WOR_MASTER及工序相關的同步作業");
            }
            else {
                log.Info("[" + WO + "]於WOR_MASTER及工序相關作業同步時，出現錯誤"); //表示SyncWoRoute當中有某段沒有成功。
            }
        }

        //負責「創建排程」的Function
        private void CreateSchedule() {

            log.Info("自動建立排程程式，執行開始!");
            log.Info("OracleWOTableName:" + OracleWOTableName + ",SQLServerScheduleTableName:" + SQLServerScheduleTableName + ",DATA_DAY_RANGE:" + DATA_DAY_RANGE);

            log.Info("");
            log.Info("首先驗證是否需要刪除舊的排程紀錄");

            List<decimal> lstToDel = new List<decimal>();
            string[] dataKeyFields = ScheduleTableKeyFields_Date.Split(',');   //日期欄位
            string[] stringKeyFields = ScheduleTableKeyFields_String.Split(','); //字串欄位
            string[] numberKeyFields = ScheduleTableKeyFields_Number.Split(','); //數字欄位

            OracleDBUtility odb = new OracleDBUtility();
            DBUtility.DBUtility sdb = new DBUtility.DBUtility();

            var clock = new Stopwatch();
            //---計時鐘開始----
            clock.Start();
            
            #region =====掃描MES當中舊的排程資料並予以刪除======
            string sql = "select * from " + SQLServerScheduleTableName + " where datediff(day," + SQLServerScheduleTableSortField + ",getdate())<=" + DATA_DAY_RANGE.ToString() + " order by " + SQLServerScheduleTableSortField + " desc";

            log.Info("主要SQL指令1: " + sql);  //←本程式最關鍵的SQL指令1(印出來看-可直接複製來作驗證使用)。

            DataTable dtSQLData = sdb.GetReader(sql);
            log.Info("MES中的排程資料筆數為: " + dtSQLData.Rows.Count+" 筆");

            //掃描MES當中舊的排程資料並予以刪除
            for (int i = 0; i < dtSQLData.Rows.Count; i++)
            {
                //要做出串接Where的條件，所以後面先寫1=1
                sql = @"SELECT * FROM " + OracleScheduleTableName + @" WHERE 1=1";

                #region ===日期欄位的處理====
                for (int x = 0; x < dataKeyFields.Length; x++)
                {
                    //看到日期欄位為空的，直接跳過
                    if (dataKeyFields[x].Equals("")) continue;

                    //如果日期欄位型別找到後，是"DBNull"，就把SQL指令串成is null
                    if (dtSQLData.Rows[i][dataKeyFields[x]].GetType().Name == "DBNull")
                    {
                        sql += " and " + dataKeyFields[x] + " is null ";
                    }
                    else
                    {
                        //↓李華寫的版本是配合兩邊資料庫都是datetime的方式。MES資料庫ZZ_CYM_WO_SCHEDULE裡面的EST_BEGIN_DATE,EST_FINISH_DATE,INSERT_DATE這三個欄位的型態有被改過!被改成了只有時間的DATE型態。若要使用這一行，要記得更新成datetime
                        sql += " and " + dataKeyFields[x] + "=to_date('" + Convert.ToDateTime(dtSQLData.Rows[i][dataKeyFields[x]]).ToString("yyyy-MM-dd HH:mm:ss") + "','YYYY-MM-DD HH24:MI:SS')";

                        //sql += " and to_char(" + dataKeyFields[x] + ",'yyyy-MM-dd') = '" + Convert.ToDateTime(dtSQLData.Rows[i][dataKeyFields[x]]).ToString("yyyy-MM-dd")+"'" ; //若沒有把ZZ_CYM_WO_SCHEDULE裡面的三個欄位改掉，就要用這種寫法跑。
                    }
                }
                #endregion

                #region ===字串欄位的處理====
                for (int x = 0; x < stringKeyFields.Length; x++)
                {
                    //看到字串欄位為空的，直接跳過
                    if (stringKeyFields[x].Equals("")) continue;

                    //如果字串欄位型別找到後，是"DBNull"，就把SQL指令串成is null
                    if (dtSQLData.Rows[i][stringKeyFields[x]].GetType().Name == "DBNull")
                    {
                        sql += " and " + stringKeyFields[x] + " is null ";
                    }
                    else
                    {
                        sql += " and " + stringKeyFields[x] + "='" + dtSQLData.Rows[i][stringKeyFields[x]].ToString() + "'";
                    }
                }
                #endregion
                
                #region ===數字欄位的處理====
                for (int x = 0; x < numberKeyFields.Length; x++)
                {
                    //看到數字欄位為空的，直接跳過
                    if (numberKeyFields[x].Equals("")) continue;

                    //如果數字欄位型別找到後，是"DBNull"，就把SQL指令串成is null
                    if (dtSQLData.Rows[i][numberKeyFields[x]].GetType().Name == "DBNull")
                    {
                        sql += " and " + numberKeyFields[x] + " is null ";
                    }
                    else
                    {
                        sql += " and " + numberKeyFields[x] + "=" + dtSQLData.Rows[i][numberKeyFields[x]].ToString() + "";
                    }
                }
                #endregion
                
                //把串出來的SQL指令，拿回Oracle找，若找不到，就可以準備把SQLServer裡面的這一筆給刪除了。
                DataTable dtOracleCheck = odb.OracleGetDataTable(sql);
                if (dtOracleCheck == null || dtOracleCheck.Rows == null || dtOracleCheck.Rows.Count == 0)
                {
                    //log.Info("查不到內容的SQL:"+ sql); //←印這一行SQL可以找出為什麼走進這個if條件的原因，若沒有問題，不用開這行。
                    log.Info("將被移除的舊排程: [工單]" + dtSQLData.Rows[i]["PRODUCE_NO"].ToString() +" [工序]"+ dtSQLData.Rows[i]["SEQNO"].ToString()+" [機台]"+ dtSQLData.Rows[i]["MACHINE_NO"].ToString()+" [寫入日期]"+ dtSQLData.Rows[i]["INSERT_DATE"].ToString().Substring(0, 10));
                    lstToDel.Add(Convert.ToDecimal(dtSQLData.Rows[i]["SID"].ToString())); //加入刪除清單中
                }

            }

            int delCount = 0;

            //把刪除清單中的紀錄刪除。
            for (int i = 0; i < lstToDel.Count; i++)
            {
                sql = "delete from " + SQLServerScheduleTableName + " where SID=" + lstToDel[i].ToString();
                sdb.doExecute(sql);
                delCount++;
            }
            log.Info("刪除了:" + delCount +"筆，舊的排程資料。");
            log.Info("");
            #endregion ==================================
            
            #region =====逐筆比對CMAS外掛中的排程資料，看是要Insert、還是update=======
            sql = @"SELECT " + OracleScheduleTableFields + " FROM " + OracleScheduleTableName + " WHERE ROUND(TO_NUMBER(sysdate - " + OracleScheduleTableSortField + @"))<=" + DATA_DAY_RANGE.ToString() + " AND SCHEDULE_DATE IS NOT NULL order by " + OracleScheduleTableSortField + " desc";

            log.Info("再來是逐一筆比對CMAS外掛中的排程資料，看是要Insert、還是update ");
            log.Info("主要SQL指令2: " + sql);  //←本程式最關鍵的SQL指令2(印出來看-可直接複製來作驗證使用)。


            int needInsertCount = 0; //計算要更新的筆數
            int needUpdateCount = 0; //計算要寫入的筆數


            DataTable dtSchedule = odb.OracleGetDataTable(sql);
            log.Info("CMAS外掛中的排程資料筆數為: " + dtSchedule.Rows.Count+" 筆");  

            for (int i = 0; i < dtSchedule.Rows.Count; i++)
            {
                //要做出串接Where的條件，所以後面先寫1=1
                sql = "select * from " + SQLServerScheduleTableName + " where 1=1 ";

                #region ===日期欄位的處理====
                for (int x = 0; x < dataKeyFields.Length; x++)
                {
                    //看到日期欄位為空的，直接跳過
                    if (dataKeyFields[x].Equals("")) continue;

                    //如果日期欄位型別找到後，是"DBNull"，就把SQL指令串成is null
                    if (dtSchedule.Rows[i][dataKeyFields[x]].GetType().Name == "DBNull")
                    {
                        sql += " and " + dataKeyFields[x] + " is null ";
                    }
                    else
                    {
                        //↓李華寫的版本是配合兩邊資料庫都是datetime的方式。MES資料庫ZZ_CYM_WO_SCHEDULE裡面的EST_BEGIN_DATE,EST_FINISH_DATE,INSERT_DATE這三個欄位的型態有被改過!被改成了只有時間的DATE型態。若要使用這一行，要記得更新成datetime
                        sql += " and CONVERT(varchar(19), " + dataKeyFields[x] + ", 120)='" + Convert.ToDateTime(dtSchedule.Rows[i][dataKeyFields[x]]).ToString("yyyy-MM-dd HH:mm:ss") + "'";

                        //sql += " and CONVERT(varchar(19), " + dataKeyFields[x] + ", 120)='" + Convert.ToDateTime(dtSchedule.Rows[i][dataKeyFields[x]]).ToString("yyyy-MM-dd") + "'";  //若沒有把ZZ_CYM_WO_SCHEDULE裡面的三個欄位改掉，就要用這種寫法跑。
                    }
                }
                #endregion

                #region ===字串欄位的處理====
                for (int x = 0; x < stringKeyFields.Length; x++)
                {
                    //看到字串欄位為空的，直接跳過
                    if (stringKeyFields[x].Equals("")) continue;

                    //如果字串欄位型別找到後，是"DBNull"，就把SQL指令串成is null
                    if (dtSchedule.Rows[i][stringKeyFields[x]].GetType().Name == "DBNull")
                    {
                        sql += " and " + stringKeyFields[x] + " is null ";
                    }
                    else
                    {
                        sql += " and " + stringKeyFields[x] + "='" + dtSchedule.Rows[i][stringKeyFields[x]].ToString() + "' ";
                    }
                }
                #endregion

                #region ===數值欄位的處理====
                for (int x = 0; x < numberKeyFields.Length; x++)
                {
                    //看到數值欄位為空的，直接跳過
                    if (numberKeyFields[x].Equals("")) continue;

                    //如果數值欄位型別找到後，是"DBNull"，就把SQL指令串成is null
                    if (dtSchedule.Rows[i][numberKeyFields[x]].GetType().Name == "DBNull")
                    {
                        sql += " and " + numberKeyFields[x] + " is null ";
                    }
                    else
                    {
                        sql += " and " + numberKeyFields[x] + "=" + dtSchedule.Rows[i][numberKeyFields[x]].ToString();
                    }
                }
                #endregion

                //把串出來的SQL指令，拿回MES找，找不到表示還沒有，就要insert。反之，就update。
                DataTable dtSQLServer = sdb.GetReader(sql);

                #region =====CMAS外掛找到的排程紀錄insert進MES=====
                if (dtSQLServer == null || dtSQLServer.Rows == null || dtSQLServer.Rows.Count == 0)
                {
                    //要做出串接的insert指令
                    sql = "insert " + SQLServerScheduleTableName + " (" + SQLServerScheduleTableFields + ") values (";
                    sql += "dbo.GetSid(),";

                    //橫向的掃瞄Oracle資料庫的欄位
                    for (int x = 0; x < dtSchedule.Columns.Count; x++)
                    {
                        //如果欄位型態是屬於System.String
                        if (dtSchedule.Columns[x].DataType.ToString().Equals("System.String"))
                        {
                            if (dtSchedule.Rows[i][x].GetType().Name == "DBNull")
                                sql += "null,";
                            else
                                sql += "'" + dtSchedule.Rows[i][x].ToString().Replace("'","") + "',";

                        }
                        //如果欄位的型態是屬於System.DateTime
                        else if (dtSchedule.Columns[x].DataType.ToString().Equals("System.DateTime"))
                        {
                            if (dtSchedule.Rows[i][x].GetType().Name == "DBNull")
                                sql += "null,";
                            else
                                sql += "'" + Convert.ToDateTime(dtSchedule.Rows[i][x]).ToString("yyyy-MM-dd HH:mm:ss") + "',";
                        }
                        //如果欄位的型態是屬於System.Decimal
                        else if (dtSchedule.Columns[x].DataType.ToString().Equals("System.Decimal"))
                        {
                            if (dtSchedule.Rows[i][x].GetType().Name == "DBNull")
                                sql += "null,";
                            else
                                sql += "" + dtSchedule.Rows[i][x].ToString() + ",";
                        }
                    }
                    sql += "'Loader',getdate(),'Loader',getdate()";
                    sql += ")";

                    log.Info("要Insert的SQL: " + sql );
                    sdb.doExecute(sql); //insert排程至MES
                    needInsertCount++;

                }
                #endregion ============================

                #region =====CMAS外掛找到的排程紀錄update回MES=====
                //反之，要用update的方式，更新排程。
                else
                {

                    bool needToUpdate = false;
                    sql = "update " + SQLServerScheduleTableName + " set ";

                    for (int iColumnIndex = 0; iColumnIndex < dtSchedule.Columns.Count; iColumnIndex++)
                    {
                        //根據欄位中資料的型別做分類，比對新舊資料的內容是否相同。不相同就needToUpdate，並串出update的SQL指令。
                        switch (dtSchedule.Columns[iColumnIndex].DataType.ToString())
                        {
                            case "System.DateTime":
                                string oraValue = "";
                                if (dtSchedule.Rows[i][iColumnIndex].GetType().Name != "DBNull") {
                                    oraValue = Convert.ToDateTime(dtSchedule.Rows[i][iColumnIndex]).ToString("yyyy-MM-dd HH:mm:ss"); //←若MES端的日期欄位當初沒有被調整，其實原本全部的日期比對都只需要用李華這一行就好!
                                    
                                    //string colName = dtSchedule.Columns[iColumnIndex].ColumnName;
                                    ////如果DB是屬於下面三個，就只能以"yyyy-MM-dd"的方式做比對
                                    //if (colName.Equals("EST_BEGIN_DATE") || colName.Equals("EST_FINISH_DATE") || colName.Equals("INSERT_DATE"))
                                    //{
                                    //    oraValue = Convert.ToDateTime(dtSchedule.Rows[i][iColumnIndex]).ToString("yyyy-MM-dd"); //不改成只看日期的話，比對起來永遠對不上!
                                    //}
                                    //else {
                                    //    oraValue = Convert.ToDateTime(dtSchedule.Rows[i][iColumnIndex]).ToString("yyyy-MM-dd HH:mm:ss"); //←若MES端的日期欄位當初沒有被調整，其實原本全部的日期比對都只需要用李華這一行就好!
                                    //}
                                    
                                }
                                string sqlValue = "";
                                if (dtSQLServer.Rows[0][dtSchedule.Columns[iColumnIndex].ColumnName].GetType().Name != "DBNull") {
                                    sqlValue = Convert.ToDateTime(dtSQLServer.Rows[0][dtSchedule.Columns[iColumnIndex].ColumnName]).ToString("yyyy-MM-dd HH:mm:ss"); //←若MES端的日期欄位當初沒有被調整，其實原本全部的日期比對都只需要用李華這一行就好!
                                    
                                    //string colName = dtSchedule.Columns[iColumnIndex].ColumnName;
                                    ////如果DB是屬於下面三個，就只能以"yyyy-MM-dd"的方式做比對
                                    //if (colName.Equals("EST_BEGIN_DATE") || colName.Equals("EST_FINISH_DATE") || colName.Equals("INSERT_DATE"))
                                    //{
                                    //    sqlValue = Convert.ToDateTime(dtSQLServer.Rows[0][dtSchedule.Columns[iColumnIndex].ColumnName]).ToString("yyyy-MM-dd"); //不改成只看日期的話，比對起來永遠對不上!
                                    //}
                                    //else {
                                    //    sqlValue = Convert.ToDateTime(dtSQLServer.Rows[0][dtSchedule.Columns[iColumnIndex].ColumnName]).ToString("yyyy-MM-dd HH:mm:ss"); //←若MES端的日期欄位當初沒有被調整，其實原本全部的日期比對都只需要用李華這一行就好!
                                    //}
                                    
                                }
                                if (!oraValue.Equals(sqlValue))
                                {
                                    log.Info("CMAS外掛當中的內容:" + oraValue + "    MES當中的內容:" + sqlValue); //若有進到這個if，表示有差異了，要做更新所以印出內容給使用這看

                                    needToUpdate = true;
                                    if (oraValue.Equals(""))
                                        sql += dtSchedule.Columns[iColumnIndex].ColumnName + "= null ,";
                                    else
                                        sql += dtSchedule.Columns[iColumnIndex].ColumnName + "= '" + oraValue + "' ,";
                                }
                                break;
                            case "System.Decimal":

                                string colName = dtSchedule.Columns[iColumnIndex].ColumnName;
                                    decimal oraDec = 0;
                                    decimal sqlDec = 0;
                                //如果欄位是下面兩種，就要以例外的方式，特別處理。(字串轉成decimal後再進行比對，這樣比對起來才會正確!)
                                if (colName.Equals("EST_USE_TIME") || colName.Equals("QTY"))
                                {
                                    oraValue = "0"; //初始時，要用字串的0才不會因為碰到null而發生問題。
                                    if (dtSchedule.Rows[i][iColumnIndex].GetType().Name != "DBNull")
                                    {
                                        oraValue = dtSchedule.Rows[i][iColumnIndex].ToString();
                                        oraDec = Decimal.Parse(oraValue);
                                    }
                                    sqlValue = "0";//初始時，要用字串的0才不會因為碰到null而發生問題。
                                    if (dtSQLServer.Rows[0][dtSchedule.Columns[iColumnIndex].ColumnName].GetType().Name != "DBNull")
                                    {
                                        sqlValue = dtSQLServer.Rows[0][dtSchedule.Columns[iColumnIndex].ColumnName].ToString();
                                        sqlDec = Decimal.Parse(sqlValue);
                                    }
                                    //用decimal做比對
                                    if (oraDec != sqlDec)
                                    {
                                        log.Info("CMAS外掛當中的內容:" + oraValue + "    MES當中的內容:" + sqlValue); //若有進到這個if，表示有差異了，要做更新所以印出內容給使用這看
                                        oraValue = oraDec.ToString(); //從decimal轉回字串

                                        needToUpdate = true;
                                        if (oraValue.Equals(""))
                                            sql += dtSchedule.Columns[iColumnIndex].ColumnName + "= null ,";
                                        else
                                            sql += dtSchedule.Columns[iColumnIndex].ColumnName + "= " + oraValue + " ,";
                                    }
                                }
                                else
                                {
                                    oraValue = "";
                                    if (dtSchedule.Rows[i][iColumnIndex].GetType().Name != "DBNull")
                                        oraValue = dtSchedule.Rows[i][iColumnIndex].ToString();
                                    sqlValue = "";
                                    if (dtSQLServer.Rows[0][dtSchedule.Columns[iColumnIndex].ColumnName].GetType().Name != "DBNull")
                                        sqlValue = dtSQLServer.Rows[0][dtSchedule.Columns[iColumnIndex].ColumnName].ToString();
                                    if (!oraValue.Equals(sqlValue))
                                    {
                                        log.Info("CMAS外掛當中的內容:" + oraValue + "    MES當中的內容:" + sqlValue); //若有進到這個if，表示有差異了，要做更新所以印出內容給使用這看

                                        needToUpdate = true;
                                        if (oraValue.Equals(""))
                                            sql += dtSchedule.Columns[iColumnIndex].ColumnName + "= null ,";
                                        else
                                            sql += dtSchedule.Columns[iColumnIndex].ColumnName + "= " + oraValue + " ,";
                                    }
                                }                                
                                break;
                            default:
                                oraValue = "";
                                if (dtSchedule.Rows[i][iColumnIndex].GetType().Name != "DBNull")
                                    oraValue = dtSchedule.Rows[i][iColumnIndex].ToString().Replace("'", ""); ;
                                sqlValue = "";
                                if (dtSQLServer.Rows[0][dtSchedule.Columns[iColumnIndex].ColumnName].GetType().Name != "DBNull")
                                    sqlValue = dtSQLServer.Rows[0][dtSchedule.Columns[iColumnIndex].ColumnName].ToString().Replace("'","");
                                if (!oraValue.Equals(sqlValue))
                                {
                                    needToUpdate = true;
                                    if (oraValue.Equals(""))
                                        sql += dtSchedule.Columns[iColumnIndex].ColumnName + "= null ,";
                                    else
                                        sql += dtSchedule.Columns[iColumnIndex].ColumnName + "= '" + oraValue + "' ,";
                                }
                                break;
                        }


                    }

                    //如果結果是需要update的，會在這裡執行。
                    if (needToUpdate)
                    {
                        if (sql.EndsWith(","))
                            sql = sql.Substring(0, sql.Length - 1);

                        string SID = dtSQLServer.Rows[0]["SID"].ToString();
                        sql += " where SID=" + SID;

                        log.Info("更新至" + SQLServerScheduleTableName+ "的SID為:"+ SID);
                        log.Info("要Update的SQL: " + sql );
                        sdb.doExecute(sql);  //update排程至MES
                        needUpdateCount++;
                    }


                }
                #endregion ============================

            }
            #endregion ==============================
            
            clock.Stop(); //創建排程的計時停止
            log.Info("");
            log.Info("本次段落，耗用的時間為:" + clock.Elapsed.TotalSeconds.ToString() + "秒");
            log.Info("更新(update)的排程資料筆數為: " + needUpdateCount + " 筆");
            log.Info("首次寫入(insert)的資料筆數為: " + needInsertCount + " 筆");

        }
        

    }
}
