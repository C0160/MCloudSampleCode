using com.digiwin.Mobile.MCloud.TransferLibrary;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Web.Script.Serialization;
using System.Xml.Linq;

namespace HANBELL
{
    public class HANBELL01
    {
        #region 頁面初使化、加班單身分頁、搜尋
        public string Get_HANBELL01_BasicSetting(XDocument pM2Pxml)
        {
            string tP2Mxml = string.Empty;//回傳值
            ProductToMCloudBuilder tP2MObject = new ProductToMCloudBuilder(null, null, null); //參數:要顯示的訊息、結果、Title//建構p2m

            try
            {
                #region 設定參數                        
                JavaScriptSerializer tSerObject = new JavaScriptSerializer();  //設定JSON轉換物件
                Dictionary<string, string> tparam = new Dictionary<string, string>();//叫用API的參數
                string tfunctionName = string.Empty;    //叫用API的方法名稱(自定義)
                string tErrorMsg = string.Empty;        //檢查是否錯誤
                #endregion

                #region 設定控件
                RCPControl tRCPbtnSave = new RCPControl("btnSave", null, null, null);       //送單
                RCPControl tRCPAdd = new RCPControl("Add", null, null, null);               //新增
                RCPControl tRCPDel = new RCPControl("Del", null, null, null);               //刪除
                RCPControl tRCPformType = new RCPControl("formType", "true", null, null);   //加班類別
                RCPControl tRCPcompany = new RCPControl("company", "true", null, null);     //公司別
                RCPControl tRCPdate = new RCPControl("date", "false", null, null);          //申請日
                RCPControl tRCPuserid = new RCPControl("userid", "false", null, null);      //申請人
                RCPControl tRCPdeptno = new RCPControl("deptno", "true", null, null);       //申請部門
                RCPControl tRCPBodyInfo = new RCPControl("BodyInfo", "true", null, null);   //申請單單身
                RCPControl tRCPDocData = new RCPControl("DocData", null, null, null);       //加班單暫存資料
                #endregion

                #region 取得畫面資料
                string tStrUserID = DataTransferManager.GetUserMappingValue(pM2Pxml, "HANBELL");    //登入者帳號
                string tStrDocData = DataTransferManager.GetControlsValue(pM2Pxml, "DocData");      //加班單暫存資料
                string tSearch = DataTransferManager.GetControlsValue(pM2Pxml, "SearchCondition");  //單身搜尋條件
                string tStrLanguage = DataTransferManager.GetDataValue(pM2Pxml, "Language");        //語系
                #endregion

                #region 處裡畫面資料
                //檢查是否已有單據資料
                if (string.IsNullOrEmpty(tStrDocData))//沒單據資料
                {
                    #region 取得預設資料
                    //設定參數
                    tfunctionName = "GetEmployeeDetail";    //取得員工明細資料
                    tparam.Clear();
                    tparam.Add("UserID", tStrUserID);       //查詢對象
                    new CustomLogger.Logger(pM2Pxml).WriteInfo("functionName : " + tfunctionName);
                    new CustomLogger.Logger(pM2Pxml).WriteInfo("Search UserID : " + tStrUserID);

                    //叫用服務
                    string tResponse = CallAPI(pM2Pxml, tfunctionName, tparam, out tErrorMsg);
                    if (!string.IsNullOrEmpty(tErrorMsg)) return ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);

                    //回傳的JSON轉Class
                    GetEmployeeDetail EmployDetail = tSerObject.Deserialize<GetEmployeeDetail>(tResponse);  //員工明細

                    //取值
                    string tCompany = string.Format("{0}-{1}", EmployDetail.company, ""/*EmployDetail.companyName*/);   //公司別代號-公司別名稱
                    string tUser = string.Format("{0}-{1}", EmployDetail.id, EmployDetail.userName);                    //申請人代號-申請人名稱
                    string tDetp = string.Format("{0}-{1}", EmployDetail.deptno, EmployDetail.deptname);                //部門代號-部門名稱
                    string tDate = DateTime.Now.ToString("yyyy/MM/dd");                                                 //申請時間
                    string tStrformType = DataTransferManager.GetControlsValue(pM2Pxml, "formType");                    //加班類別code
                    string fStrormTypeDesc = "1".Equals(tStrformType) ? "平日加班" : "2".Equals(tStrformType) ? "雙休加班" : "節日加班";    //加班類別name

                    #region 取得公司別名稱(如果GetEmployeeDetail有回傳名稱的話可以省掉)
                    tfunctionName = "GetCompany";    //取得員工明細資料
                    tparam.Clear();
                    new CustomLogger.Logger(pM2Pxml).WriteInfo("functionName : " + tfunctionName);
                    new CustomLogger.Logger(pM2Pxml).WriteInfo("Search UserID : " + tStrUserID);

                    //取公司別值
                    tResponse = CallAPI(pM2Pxml, tfunctionName, tparam, out tErrorMsg);
                    if (!string.IsNullOrEmpty(tErrorMsg)) return ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);

                    tResponse = "{\"CompanyItems\" : " + tResponse + "}";                                           //為符合class格式自己加上去的
                    GetCompany CompanyClass = tSerObject.Deserialize<GetCompany>(tResponse);                        //轉class方便取用
                    DataTable tCompanyTable = CompanyClass.GetCompanyTable(EmployDetail.company);                   //取公司別table
                    DataRow[] tCompanyRow = tCompanyTable.Select("company = '" + EmployDetail.company + "'");       //查詢指定company
                    if (tCompanyRow.Count() > 0) tCompany = tCompanyRow[0]["company"] + "-" + tCompanyRow[0]["name"];
                    #endregion

                    #endregion

                    #region 給值
                    //設定加班單資料
                    Docdata DocDataClass = new Docdata();
                    DocDataClass.head = new HeadData(tCompany, tDate, tUser, tDetp, tStrformType, fStrormTypeDesc);
                    string DocDataStr = tSerObject.Serialize(DocDataClass);//Class轉成JSON

                    //給定畫面欄位值
                    tRCPcompany.Value = tCompany;       //公司別
                    tRCPdate.Value = tDate;             //申請日
                    tRCPuserid.Value = tUser;           //申請人
                    tRCPdeptno.Value = tDetp;           //申請部門
                    tRCPDocData.Value = DocDataStr;     //加班單暫存資料
                    #endregion
                }
                else//有單據資料
                {
                    #region 顯示加班單資料
                    //先將暫存資料轉成Class方便取用
                    Docdata DocDataClass = tSerObject.Deserialize<Docdata>(tStrDocData);//加班單暫存資料

                    //設定畫面資料
                    tRCPcompany.Value = DocDataClass.head.company;      //公司別
                    tRCPdate.Value = DocDataClass.head.date;            //申請日
                    tRCPuserid.Value = DocDataClass.head.id;            //申請人
                    tRCPdeptno.Value = DocDataClass.head.deptno;        //申請部門
                    tRCPDocData.Value = tStrDocData;                    //加班單暫存資料

                    //因實際不影響單身新增資料, 故取消唯讀控制
                    ////檢查是否有單身資料
                    //if (!(DocDataClass.body.Count > 0))//沒有單身資料
                    //{
                    //    //設定單頭欄位屬性
                    //    tRCPcompany.Enable = "true";   //公司別
                    //    tRCPdeptno.Enable = "true";    //申請部門
                    //}
                    //else//有單身資料
                    //{
                    //    //設定單頭屬性
                    //    tRCPcompany.Enable = "false";   //公司別
                    //    tRCPdeptno.Enable = "false";    //申請部門
                    //}

                    //整理Table轉成符合顯示的格式
                    DataTable tBodyTable = DocDataClass.GetBodyDataTable(tSearch);
                    tBodyTable.Columns.Add("dateC");//日期,有加/
                    tBodyTable.Columns.Add("time"); //加班時間
                    tBodyTable.Columns.Add("food"); //供餐
                    for (int i = 0; i < tBodyTable.Rows.Count; i++)
                    {
                        string tstarttime = tBodyTable.Rows[i]["starttime"].ToString().Insert(2, ":");
                        string tdate = tBodyTable.Rows[i]["date"].ToString().Insert(4, "/").Insert(7, "/");
                        string tendtime = tBodyTable.Rows[i]["endtime"].ToString().Insert(2, ":");
                        string tlunch = tBodyTable.Rows[i]["lunch"].ToString();
                        string tdinner = tBodyTable.Rows[i]["dinner"].ToString();

                        tBodyTable.Rows[i]["dateC"] = tdate;
                        tBodyTable.Rows[i]["time"] = tstarttime + "-" + tendtime;
                        tBodyTable.Rows[i]["food"] = ("N".Equals(tlunch) && "N".Equals(tdinner)) ? "無" :
                                                     ("Y".Equals(tlunch) && "Y".Equals(tdinner)) ? "午、晚餐" :
                                                     ("Y".Equals(tlunch) && "N".Equals(tdinner)) ? "午餐" : "晚餐";

                    }

                    //設定單身資料(含分頁)
                    SetRCPControlpage(pM2Pxml, tBodyTable, ref tRCPBodyInfo);
                    #endregion
                }
                #endregion

                #region 設定物件屬性
                //清除GridView勾選
                tRCPBodyInfo.Value = "";

                //單身清單屬性
                tRCPBodyInfo.AddCustomTag("DesplayColumn", "id§dateC§time§worktime§food§note");                
                tRCPBodyInfo.AddCustomTag("StructureStyle", "id:R1:C1:3:::§dateC:R1:C4:3:::§time:R2:C1:3:::§worktime:R2:C4:1:::§food:R2:C5:2:::§note:R3:C1:6:::");

                //設定多語系
                if ("Lang01".Equals(tStrLanguage))
                {
                    //繁體
                    tRCPbtnSave.Title = "送單";
                    tRCPAdd.Title = "新增";
                    tRCPDel.Title = "刪除";
                    tRCPformType.Title = "加班類別";
                    tRCPcompany.Title = "公司別";
                    tRCPdate.Title = "申請日";
                    tRCPuserid.Title = "申請人";
                    tRCPdeptno.Title = "申請部門";
                    tRCPBodyInfo.AddCustomTag("ColumnsName", "加班人員§加班日期§加班時間§時數§供餐§加班內容");
                }
                else
                {
                    //簡體
                    tRCPbtnSave.Text = "送单";
                    tRCPAdd.Text = "新增";
                    tRCPDel.Text = "删除";
                    tRCPformType.Title = "加班类别";
                    tRCPcompany.Title = "公司别";
                    tRCPdate.Title = "申请日";
                    tRCPuserid.Title = "申请人";
                    tRCPdeptno.Title = "申请部门";
                    tRCPBodyInfo.AddCustomTag("ColumnsName", "加班人员§加班日期§加班时间§时数§供餐§加班内容");
                }
                #endregion

                //處理回傳
                tP2MObject.AddRCPControls(tRCPformType, tRCPcompany, tRCPdate, tRCPuserid, tRCPdeptno, tRCPBodyInfo, tRCPDocData);
            }
            catch (Exception err)
            {
                ReturnErrorMsg(pM2Pxml, ref tP2MObject, "Get_HANBELL01_BasicSetting Error : " + err.Message.ToString());
            }

            tP2Mxml = tP2MObject.ToDucument().ToString();
            return tP2Mxml;            
        }
        #endregion

        #region 加班類別異動
        public string Get_HANBELL01_formType_OnBuler(XDocument pM2Pxml)
        {
            string tP2Mxml = string.Empty;//回傳值
            ProductToMCloudBuilder tP2MObject = new ProductToMCloudBuilder(null, null, null); //參數:要顯示的訊息、結果、Title//建構p2m

            try
            {
                #region 設定參數                        
                JavaScriptSerializer tSerObject = new JavaScriptSerializer();  //設定JSON轉換物件
                Dictionary<string, string> tparam = new Dictionary<string, string>();//叫用API的參數
                string tfunctionName = string.Empty;    //叫用API的方法名稱(自定義)
                string tErrorMsg = string.Empty;        //檢查是否錯誤
                #endregion

                #region 設定控件
                RCPControl tRCPDocData = new RCPControl("DocData", null, null, null);       //加班單據隱藏欄位
                #endregion

                #region 取得畫面資料
                string tStrDocData = DataTransferManager.GetControlsValue(pM2Pxml, "DocData");                      //加班單據暫存
                string tStrformType = DataTransferManager.GetControlsValue(pM2Pxml, "formType");                    //加班類別code
                string fStrormTypeDesc = "1".Equals(tStrformType) ? "平日加班" : "2".Equals(tStrformType) ? "雙休加班" : "節日加班";    //加班類別name
                #endregion

                #region 處理畫面資料
                //先將暫存資料轉成Class方便取用
                Docdata DocDataClass = tSerObject.Deserialize<Docdata>(tStrDocData);//加班單暫存資料

                //修改公司別資料
                DocDataClass.head.formType = tStrformType;
                DocDataClass.head.formTypeDesc = fStrormTypeDesc;

                //加班單Class轉成JSON
                string tDocdataJSON = tSerObject.Serialize(DocDataClass);

                //給值
                tRCPDocData.Value = tDocdataJSON;
                #endregion

                //處理回傳
                tP2MObject.AddRCPControls(tRCPDocData);
            }
            catch (Exception err)
            {
                ReturnErrorMsg(pM2Pxml, ref tP2MObject, "Get_HANBELL01_formType_OnBuler Error : " + err.Message.ToString());
            }

            tP2Mxml = tP2MObject.ToDucument().ToString();
            return tP2Mxml;
        }
        #endregion

        #region 公司別開窗
        public string Get_HANBELL01_Company_OP(XDocument pM2Pxml)
        {
            string tP2Mxml = string.Empty;//回傳值
            ProductToMCloudBuilder tP2MObject = new ProductToMCloudBuilder(null, null, null); //參數:要顯示的訊息、結果、Title//建構p2m

            try
            {
                #region 設定參數                        
                JavaScriptSerializer tSerObject = new JavaScriptSerializer();  //設定JSON轉換物件
                Dictionary<string, string> tparam = new Dictionary<string, string>();//叫用API的參數
                string tfunctionName = string.Empty;    //叫用API的方法名稱(自定義)
                string tErrorMsg = string.Empty;        //檢查是否錯誤
                #endregion

                #region 設定控件
                RDControl tRDcompany = new RDControl(new DataTable());  //公司別清單
                #endregion

                #region 取得畫面資料
                string tSearch = DataTransferManager.GetControlsValue(pM2Pxml, "SearchCondition");  //單身搜尋條件
                string tStrLanguage = DataTransferManager.GetDataValue(pM2Pxml, "Language");        //語系
                #endregion

                #region 處理畫面資料
                #region 取得公司別清單資料
                //設定參數
                tfunctionName = "GetCompany";    //取得公司別
                tparam.Clear();
                new CustomLogger.Logger(pM2Pxml).WriteInfo("functionName : " + tfunctionName);

                //叫用服務
                string tResponse = CallAPI(pM2Pxml, tfunctionName, tparam, out tErrorMsg);
                if (!string.IsNullOrEmpty(tErrorMsg)) return ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);

                tResponse = "{\"CompanyItems\" : " + tResponse + "}";                       //為符合class格式自己加上去的
                GetCompany CompanyClass = tSerObject.Deserialize<GetCompany>(tResponse);    //轉class方便取用
                #endregion

                #region 給值
                //取得公司別Table
                DataTable tCompanyTable = CompanyClass.GetCompanyTable(tSearch);
                tCompanyTable.Columns.Add("companyC");//    Control id + C = 開窗控件要顯示的值
                for (int i = 0; i < tCompanyTable.Rows.Count; i++) { tCompanyTable.Rows[i]["companyC"] = tCompanyTable.Rows[i]["company"] + "-" + tCompanyTable.Rows[i]["name"]; }//設定顯示的值

                //設定公司別清單資料(含分頁)
                SetRDControlpage(pM2Pxml, tCompanyTable, ref tRDcompany);
                #endregion
                #endregion

                #region 設定物件屬性
                //公司別清單屬性
                tRDcompany.AddCustomTag("DisplayColumn", "company§name§companyC");
                tRDcompany.AddCustomTag("StructureStyle", "company:R1:C1:1§name:R1:C2:1");

                //設定多語系
                if ("Lang01".Equals(tStrLanguage))
                {
                    //繁體
                    tRDcompany.AddCustomTag("ColumnsName", "代號§名稱§公司別");
                }
                else
                {
                    //簡體
                    tRDcompany.AddCustomTag("ColumnsName", "代号§名称§公司别");
                }
                #endregion

                //處理回傳
                tP2MObject.AddRDControl(tRDcompany);

            }
            catch (Exception err)
            {
                ReturnErrorMsg(pM2Pxml, ref tP2MObject, "Get_HANBELL01_Company_OP Error : " + err.Message.ToString());
            }

            tP2Mxml = tP2MObject.ToDucument().ToString();
            return tP2Mxml;
        }
        #endregion

        #region 公司別異動
        public string Get_HANBELL01_company_OnBlur(XDocument pM2Pxml)
        {
            string tP2Mxml = string.Empty;//回傳值
            ProductToMCloudBuilder tP2MObject = new ProductToMCloudBuilder(null, null, null); //參數:要顯示的訊息、結果、Title//建構p2m

            try
            {
                #region 設定參數                        
                JavaScriptSerializer tSerObject = new JavaScriptSerializer();  //設定JSON轉換物件
                Dictionary<string, string> tparam = new Dictionary<string, string>();//叫用API的參數
                string tfunctionName = string.Empty;    //叫用API的方法名稱(自定義)
                string tErrorMsg = string.Empty;        //檢查是否錯誤
                #endregion

                #region 設定控件
                RCPControl tRCPDocData = new RCPControl("DocData", null, null, null);       //加班單據隱藏欄位
                #endregion

                #region 取得畫面資料
                string tStrcompany = DataTransferManager.GetControlsValue(pM2Pxml, "companyC");  //公司別(C是取外顯值)
                string tStrDocData = DataTransferManager.GetControlsValue(pM2Pxml, "DocData");  //加班單據暫存
                #endregion

                #region 處理畫面資料
                //先將暫存資料轉成Class方便取用
                Docdata DocDataClass = tSerObject.Deserialize<Docdata>(tStrDocData);//加班單暫存資料

                //修改公司別資料
                DocDataClass.head.company = tStrcompany;

                //加班單Class轉成JSON
                string tDocdataJSON = tSerObject.Serialize(DocDataClass);

                //給值
                tRCPDocData.Value = tDocdataJSON;
                #endregion

                //處理回傳
                tP2MObject.AddRCPControls(tRCPDocData);
            }
            catch (Exception err)
            {
                ReturnErrorMsg(pM2Pxml, ref tP2MObject, "Get_HANBELL01_Company_OP_OnBlur Error : " + err.Message.ToString());
            }

            tP2Mxml = tP2MObject.ToDucument().ToString();
            return tP2Mxml;
        }
        #endregion

        #region 部門開窗
        public string Get_HANBELL01_deptno_OP(XDocument pM2Pxml)
        {
            string tP2Mxml = string.Empty;//回傳值
            ProductToMCloudBuilder tP2MObject = new ProductToMCloudBuilder(null, null, null); //參數:要顯示的訊息、結果、Title//建構p2m

            try
            {
                #region 設定參數                        
                JavaScriptSerializer tSerObject = new JavaScriptSerializer();  //設定JSON轉換物件
                Dictionary<string, string> tparam = new Dictionary<string, string>();//叫用API的參數
                string tfunctionName = string.Empty;    //叫用API的方法名稱(自定義)
                string tErrorMsg = string.Empty;        //檢查是否錯誤
                #endregion

                #region 設定控件
                RDControl tRDdeptno = new RDControl(new DataTable());  //部門清單
                #endregion

                #region 取得畫面資料
                string tStrUserID = DataTransferManager.GetUserMappingValue(pM2Pxml, "HANBELL");    //登入者帳號
                string tSearch = DataTransferManager.GetControlsValue(pM2Pxml, "SearchCondition");  //單身搜尋條件
                string tStrLanguage = DataTransferManager.GetDataValue(pM2Pxml, "Language");        //語系
                #endregion

                #region 處理畫面資料
                #region 取得公司別清單資料
                //設定參數
                tfunctionName = "GetEmployeeDepartment";    //查詢員工部門
                tparam.Clear();
                tparam.Add("userId", tStrUserID);           //查詢對象
                new CustomLogger.Logger(pM2Pxml).WriteInfo("functionName : " + tfunctionName);

                //叫用服務
                string tResponse = CallAPI(pM2Pxml, tfunctionName, tparam, out tErrorMsg);
                if (!string.IsNullOrEmpty(tErrorMsg)) return ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);

                tResponse = "{\"EmployeeDepartmentItems\" : " + tResponse + "}";                                    //為符合class格式自己加上去的
                GetEmployeeDepartment DeparmentClass = tSerObject.Deserialize<GetEmployeeDepartment>(tResponse);    //轉class方便取用
                #endregion

                #region 給值
                //取得部門Table
                DataTable tDeparmentTable = DeparmentClass.GetDeparmentTable(tSearch);

                //轉換成符合格式的table
                if (tDeparmentTable.Columns.Contains("id")) { tDeparmentTable.Columns["id"].ColumnName = "deptno"; }
                if (tDeparmentTable.Columns.Contains("organizationUnitName")) { tDeparmentTable.Columns["organizationUnitName"].ColumnName = "deptnoName"; }
                tDeparmentTable.Columns.Add("deptnoC");
                for(int i=0;i< tDeparmentTable.Rows.Count; i++) { tDeparmentTable.Rows[i]["deptnoC"] = tDeparmentTable.Rows[i]["deptno"] + "-" + tDeparmentTable.Rows[i]["deptnoName"]; }

                //設定公司別清單資料(含分頁)
                SetRDControlpage(pM2Pxml, tDeparmentTable, ref tRDdeptno);
                #endregion
                #endregion

                #region 設定物件屬性
                //公司別清單屬性
                tRDdeptno.AddCustomTag("DisplayColumn", "deptno§deptnoName§deptnoC");
                tRDdeptno.AddCustomTag("ColumnsName", "代號§名稱§部門");
                tRDdeptno.AddCustomTag("StructureStyle", "deptno:R1:C1:1§deptnoName:R1:C2:1");

                //設定多語系
                if ("Lang01".Equals(tStrLanguage))
                {
                    //繁體
                    tRDdeptno.AddCustomTag("ColumnsName", "代號§名稱§部門");
                }
                else
                {
                    //簡體
                    tRDdeptno.AddCustomTag("ColumnsName", "代号§名称§部门");
                }
                #endregion

                //處理回傳
                tP2MObject.AddRDControl(tRDdeptno);

            }
            catch (Exception err)
            {
                ReturnErrorMsg(pM2Pxml, ref tP2MObject, "Get_HANBELL01_deptno_OP Error : " + err.Message.ToString());
            }

            tP2Mxml = tP2MObject.ToDucument().ToString();
            return tP2Mxml;
        }
        #endregion

        #region 部門異動
        public string Get_HANBELL01_deptno_OnBlur(XDocument pM2Pxml)
        {
            string tP2Mxml = string.Empty;//回傳值
            ProductToMCloudBuilder tP2MObject = new ProductToMCloudBuilder(null, null, null); //參數:要顯示的訊息、結果、Title//建構p2m

            try
            {
                #region 設定參數                        
                JavaScriptSerializer tSerObject = new JavaScriptSerializer();  //設定JSON轉換物件
                Dictionary<string, string> tparam = new Dictionary<string, string>();//叫用API的參數
                string tfunctionName = string.Empty;    //叫用API的方法名稱(自定義)
                string tErrorMsg = string.Empty;        //檢查是否錯誤
                #endregion

                #region 設定控件
                RCPControl tRCPDocData = new RCPControl("DocData", null, null, null);       //加班單據隱藏欄位
                #endregion

                #region 取得畫面資料
                string tStrdeptno = DataTransferManager.GetControlsValue(pM2Pxml, "deptnoC");    //申請部門(C是取外顯值)
                string tStrDocData = DataTransferManager.GetControlsValue(pM2Pxml, "DocData");  //加班單據暫存
                #endregion

                #region 處理畫面資料
                //先將暫存資料轉成Class方便取用
                Docdata DocDataClass = tSerObject.Deserialize<Docdata>(tStrDocData);//加班單暫存資料

                //修改部門資料
                DocDataClass.head.deptno = tStrdeptno;

                //加班單Class轉成JSON
                string tDocdataJSON = tSerObject.Serialize(DocDataClass);

                //給值
                tRCPDocData.Value = tDocdataJSON;
                #endregion

                //處理回傳
                tP2MObject.AddRCPControls(tRCPDocData);
            }
            catch (Exception err)
            {
                ReturnErrorMsg(pM2Pxml, ref tP2MObject, "Get_HANBELL01_deptno_OnBlur Error : " + err.Message.ToString());
            }

            tP2Mxml = tP2MObject.ToDucument().ToString();
            return tP2Mxml;
        }
        #endregion

        #region 刪除單身
        public string Get_HANBELL01_Del(XDocument pM2Pxml)
        {
            string tP2Mxml = string.Empty;//回傳值
            ProductToMCloudBuilder tP2MObject = new ProductToMCloudBuilder(null, null, null); //參數:要顯示的訊息、結果、Title//建構p2m

            try
            {
                #region 設定參數                        
                JavaScriptSerializer tSerObject = new JavaScriptSerializer();  //設定JSON轉換物件
                string tErrorMsg = string.Empty;        //檢查是否錯誤
                #endregion

                #region 設定控件
                RCPControl tRCPcompany = new RCPControl("company", "true", null, null);     //公司別
                RCPControl tRCPdeptno = new RCPControl("deptno", "true", null, null);       //申請部門
                RCPControl tRCPDocData = new RCPControl("DocData", null, null, null);       //加班單據隱藏欄位
                RCPControl tRCPBodyInfo = new RCPControl("BodyInfo", "true", null, null);   //申請單單身
                #endregion

                #region 取得畫面資料
                string tStrBodyInfo = DataTransferManager.GetControlsValue(pM2Pxml, "BodyInfo");    //加班單單身勾選的資料
                string tStrDocData = DataTransferManager.GetControlsValue(pM2Pxml, "DocData");      //加班單據暫存
                string tSearch = DataTransferManager.GetControlsValue(pM2Pxml, "SearchCondition");  //單身搜尋條件
                #endregion

                if (!string.IsNullOrEmpty(tStrBodyInfo))
                {
                    #region 處理畫面資料
                    //先將暫存資料轉成Class方便取用
                    Docdata DocDataClass = tSerObject.Deserialize<Docdata>(tStrDocData);//加班單暫存資料

                    //刪除單身資料
                    foreach (var item in tStrBodyInfo.Split('§'))
                    {
                        foreach (var bodyitem in DocDataClass.body)
                        {
                            if (item.Equals(bodyitem.item))
                            {
                                DocDataClass.body.Remove(bodyitem);
                                break;
                            }
                        }
                    }                    

                    #region 顯示加班單資料                    
                    tRCPDocData.Value = tSerObject.Serialize(DocDataClass);                    //加班單暫存資料//加班單Class轉成JSON

                    //檢查是否有單身資料
                    if (!(DocDataClass.body.Count > 0))//沒有單身資料
                    {
                        //設定單頭欄位屬性
                        tRCPcompany.Enable = "true";   //公司別
                        tRCPdeptno.Enable = "true";    //申請部門
                    }
                    else//有單身資料
                    {
                        //設定單頭屬性
                        tRCPcompany.Enable = "false";   //公司別
                        tRCPdeptno.Enable = "false";    //申請部門
                    }

                    //整理Table轉成符合顯示的格式
                    DataTable tBodyTable = DocDataClass.GetBodyDataTable(tSearch);
                    tBodyTable.Columns.Add("dateC");//日期,有加/
                    tBodyTable.Columns.Add("time"); //加班時間
                    tBodyTable.Columns.Add("food"); //供餐
                    for (int i = 0; i < tBodyTable.Rows.Count; i++)
                    {
                        string tstarttime = tBodyTable.Rows[i]["starttime"].ToString().Insert(2, ":");
                        string tdate = tBodyTable.Rows[i]["date"].ToString().Insert(4, "/").Insert(7, "/");
                        string tendtime = tBodyTable.Rows[i]["endtime"].ToString().Insert(2, ":");
                        string tlunch = tBodyTable.Rows[i]["lunch"].ToString();
                        string tdinner = tBodyTable.Rows[i]["dinner"].ToString();

                        tBodyTable.Rows[i]["dateC"] = tdate;
                        tBodyTable.Rows[i]["time"] = tstarttime + "-" + tendtime;
                        tBodyTable.Rows[i]["food"] = ("N".Equals(tlunch) && "N".Equals(tdinner)) ? "無" :
                                                     ("Y".Equals(tlunch) && "Y".Equals(tdinner)) ? "午、晚餐" :
                                                     ("Y".Equals(tlunch) && "N".Equals(tdinner)) ? "午餐" : "晚餐";

                    }

                    //設定單身資料(含分頁)
                    SetRCPControlpage(pM2Pxml, tBodyTable, ref tRCPBodyInfo);
                    #endregion

                    #endregion

                    //處理回傳
                    tP2MObject.AddRCPControls(tRCPcompany, tRCPdeptno, tRCPDocData, tRCPBodyInfo);
                }
            }
            catch (Exception err)
            {
                ReturnErrorMsg(pM2Pxml, ref tP2MObject, "Get_HANBELL01_Del Error : " + err.Message.ToString());
            }

            tP2Mxml = tP2MObject.ToDucument().ToString();
            return tP2Mxml;
        }
        #endregion

        #region 立單
        public string Get_HANBELL01_CreateDoc(XDocument pM2Pxml)
        {
            string tP2Mxml = string.Empty;//回傳值
            ProductToMCloudBuilder tP2MObject = new ProductToMCloudBuilder(null, null, null); //參數:要顯示的訊息、結果、Title//建構p2m

            try
            {
                #region 設定參數                        
                JavaScriptSerializer tSerObject = new JavaScriptSerializer();  //設定JSON轉換物件
                Dictionary<string, string> tparam = new Dictionary<string, string>();//叫用API的參數
                string tfunctionName = string.Empty;    //叫用API的方法名稱(自定義)
                string tErrorMsg = string.Empty;        //檢查是否錯誤
                #endregion

                #region 設定控件
                RCPControl tRCPformType = new RCPControl("formType", "true", null, null);   //加班類別
                RCPControl tRCPcompany = new RCPControl("company", "true", null, null);     //公司別
                RCPControl tRCPdate = new RCPControl("date", "false", null, null);          //申請日
                RCPControl tRCPuserid = new RCPControl("userid", "false", null, null);      //申請人
                RCPControl tRCPdeptno = new RCPControl("deptno", "true", null, null);       //申請部門
                RCPControl tRCPBodyInfo = new RCPControl("BodyInfo", "true", null, null);   //申請單單身
                RCPControl tRCPDocData = new RCPControl("DocData", null, null, null);       //加班單據隱藏欄位
                #endregion

                #region 取得畫面資料
                string tStrcompany = DataTransferManager.GetControlsValue(pM2Pxml, "company");      //公司別
                string tStrdeptno = DataTransferManager.GetControlsValue(pM2Pxml, "deptno");        //申請部門
                string tStrUserID = DataTransferManager.GetUserMappingValue(pM2Pxml, "HANBELL");    //登入者帳號
                string tStrDocData = DataTransferManager.GetControlsValue(pM2Pxml, "DocData");      //加班單據暫存
                string tStrLanguage = DataTransferManager.GetDataValue(pM2Pxml, "Language");        //語系

                //先將暫存資料轉成Class方便取用
                Docdata DocDataClass = tSerObject.Deserialize<Docdata>(tStrDocData);                //加班單暫存資料
                #endregion

                #region 檢查錯誤
                //檢查公司別必填
                if(string.IsNullOrEmpty(tStrcompany))
                {
                    //設定多語系
                    if ("Lang01".Equals(tStrLanguage))
                    {
                        //繁體
                        tErrorMsg += "公司別必填\r\n";
                    }
                    else
                    {
                        //簡體
                        tErrorMsg += "公司别必填\r\n";
                    }
                }

                //檢查申請部門必填
                if (string.IsNullOrEmpty(tStrdeptno))
                {
                    //設定多語系
                    if ("Lang01".Equals(tStrLanguage))
                    {
                        //繁體
                        tErrorMsg += "申請部門必填\r\n";
                    }
                    else
                    {
                        //簡體
                        tErrorMsg += "申请部门必填\r\n";
                    }
                }

                //檢查單身是否存在
                if (string.IsNullOrEmpty(tStrDocData) || DocDataClass.body.Count <= 0)
                {
                    //設定多語系
                    if ("Lang01".Equals(tStrLanguage))
                    {
                        //繁體
                        tErrorMsg += "請確定單身資料是否存在";
                    }
                    else
                    {
                        //簡體
                        tErrorMsg += "请确定单身资料是否存在";
                    }
                }

                if(!string.IsNullOrEmpty(tErrorMsg)) return ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);
                #endregion 

                #region 整理Request
                #region 格式
                /*
                 http://i2.hanbell.com.cn:8080/Hanbell-JRS/api/efgp/hkgl034/create?appid=1505278334853&token=0ec858293fccfad55575e26b0ce31177

                 {
                    "head" : {
                        "company" : "C",
                        "date" : "2017/09/23",
                        "id" : "C0160",
                        "deptno" : "13120",
    　                     "formType" : "1",
    　                     "formTypeDesc" : "平日加班"
                    },
                    "body" : [{
                            "lunch" : "N",
                            "dinner" : "N",
                            "deptno" : "13120",
                            "id" : "C0160",
                            "date" : "2017/9/23",
                            "starttime" : "17:10",
                            "endtime" : "18:10",
                            "worktime" : "1",
                            "note" : "備註"
                        }, {
                            "lunch" : "N",
                            "dinner" : "N",
                            "deptno" : "13120",
                            "id" : "C0160",
                            "date" : "2017/9/24",
                            "starttime" : "17:10",
                            "endtime" : "18:10",
                            "worktime" : "1",
                            "note" : "備註"
                        }
                    ],
                    "appid" : "1505278334853",
                    "token" : "0ec858293fccfad55575e26b0ce31177"
                }


                response
                {"code":"200","msg":"PKG_HK_GL03400000066"}
                */
                #endregion

                //單頭
                DocDataClass.head.company = DocDataClass.head.company.Split('-')[0];    //公司別只取id
                DocDataClass.head.id = DocDataClass.head.id.Split('-')[0];              //申請人只取id
                DocDataClass.head.deptno = DocDataClass.head.deptno.Split('-')[0];      //申請部門只取id

                //單身
                foreach (var bodyitem in DocDataClass.body)
                {
                    bodyitem.deptno = bodyitem.deptno.Split('-')[0];                //加班部門只取id
                    bodyitem.id = bodyitem.id.Split('-')[0];                        //加班人員只取id
                    bodyitem.date = bodyitem.date.Insert(4, "/").Insert(7, "/");    //日期加/
                    bodyitem.starttime = bodyitem.starttime.Insert(2, ":");         //時間加:
                    bodyitem.endtime = bodyitem.endtime.Insert(2, ":");             //時間加:
                }

                new CustomLogger.Logger(pM2Pxml).WriteInfo("CreateDoc  Request: \r\n" + tSerObject.Serialize(DocDataClass));
                #endregion

                #region 取得Response
                //叫用API
                string uri = string.Format("{0}{1}?{2}", LoadConfig.GetWebConfig("APIURI"), "efgp/hkgl034/create", LoadConfig.GetWebConfig("APIKey"));
                string tBodyContext = tSerObject.Serialize(DocDataClass);

                ////測試用
                //string uri = "http://i2.hanbell.com.cn:8080/Hanbell-JRS/api/efgp/hkgl034/create?appid=1505278334853&token=0ec858293fccfad55575e26b0ce31177";
                //string param = "{\"head\":{\"company\":\"C\",\"date\":\"2017/09/23\",\"id\":\"C0160\",\"deptno\":\"13120\",\"formType\":\"1\",\"formTypeDesc\":\"平日加班\"},\"body\":[{\"lunch\":\"N\",\"dinner\":\"N\",\"deptno\":\"13120\",\"id\":\"C0160\",\"date\":\"2017/9/23\",\"starttime\":\"17:10\",\"endtime\":\"18:10\",\"worktime\":\"1\",\"note\":\"備註\"},{\"lunch\":\"N\",\"dinner\":\"N\",\"deptno\":\"13120\",\"id\":\"C0160\",\"date\":\"2017/9/24\",\"starttime\":\"17:10\",\"endtime\":\"18:10\",\"worktime\":\"1\",\"note\":\"備註\"}],\"appid\":\"1505278334853\",\"token\":\"0ec858293fccfad55575e26b0ce31177\"}";

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);   //建立連線
                request.Method = "POST";                                                //設定叫用模型
                request.ContentType = "application/json";
                request.Timeout = 300000;                                               //設定timeout時間
                string tResponse = string.Empty;

                //Get Response
                try
                {
                    using (var streamWriter = new StreamWriter(request.GetRequestStream()))
                    {
                        streamWriter.Write(tBodyContext);
                    }

                    using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)//送出參數取得回傳值
                    {
                        if (response.StatusCode == HttpStatusCode.OK)//判斷是否成功
                        {
                            using (Stream reponseStream = response.GetResponseStream())//取得回傳結果
                            {
                                using (StreamReader rd = new StreamReader(reponseStream, Encoding.UTF8))//讀取
                                {
                                    string responseString = rd.ReadToEnd();

                                    tResponse = responseString;
                                }
                            }
                        }
                        else
                        {
                            tErrorMsg = "叫用服務失敗:" + response.StatusCode.ToString();
                        }
                    }
                }
                catch (Exception ex)
                {
                    tErrorMsg = "叫用服務失敗:" + ex.Message.ToString();
                }

                new CustomLogger.Logger(pM2Pxml).WriteInfo("CreateDoc  Response: \r\n" + tResponse);
                #endregion

                #region 處理畫面資料
                if (!string.IsNullOrEmpty(tErrorMsg))
                {
                    ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);
                }
                else
                {
                    //轉成class 方便取用
                    CreateDocResult tCreateDocResultClass = tSerObject.Deserialize<CreateDocResult>(tResponse);

                    //判斷回傳
                    if ("200".Equals(tCreateDocResultClass.code))
                    {
                        string tMsg = string.Empty;

                        //設定多語系
                        if ("Lang01".Equals(tStrLanguage))
                        {
                            //繁體
                            tMsg = "請假單建立成功，單號 :\r\n";
                        }
                        else
                        {
                            //簡體
                            tMsg = "请假单建立成功，单号 :\r\n";
                        }

                        tP2MObject.Message = tMsg + tCreateDocResultClass.msg;
                        tP2MObject.Result = "false";

                        #region 取得預設資料
                        //設定參數
                        tfunctionName = "GetEmployeeDetail";    //取得員工明細資料
                        tparam.Clear();
                        tparam.Add("UserID", tStrUserID);       //查詢對象
                        new CustomLogger.Logger(pM2Pxml).WriteInfo("functionName : " + tfunctionName);
                        new CustomLogger.Logger(pM2Pxml).WriteInfo("Search UserID : " + tStrUserID);

                        //叫用服務
                        tResponse = CallAPI(pM2Pxml, tfunctionName, tparam, out tErrorMsg);
                        if (!string.IsNullOrEmpty(tErrorMsg)) return ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);

                        //回傳的JSON轉Class
                        GetEmployeeDetail EmployDetail = tSerObject.Deserialize<GetEmployeeDetail>(tResponse);  //員工明細

                        //取值
                        string tCompany = string.Format("{0}-{1}", EmployDetail.company, ""/*EmployDetail.companyName*/);   //公司別代號-公司別名稱
                        string tUser = string.Format("{0}-{1}", EmployDetail.id, EmployDetail.userName);                    //申請人代號-申請人名稱
                        string tDetp = string.Format("{0}-{1}", EmployDetail.deptno, EmployDetail.deptname);                //部門代號-部門名稱
                        string tDate = DateTime.Now.ToString("yyyy/MM/dd");                                                 //申請時間
                        string tStrformType = DataTransferManager.GetControlsValue(pM2Pxml, "formType");                    //加班類別code
                        string fStrormTypeDesc = "1".Equals(tStrformType) ? "平日加班" : "2".Equals(tStrformType) ? "雙休加班" : "節日加班";    //加班類別name

                        #region 取得公司別名稱(如果GetEmployeeDetail有回傳名稱的話可以省掉)
                        tfunctionName = "GetCompany";    //取得員工明細資料
                        tparam.Clear();
                        new CustomLogger.Logger(pM2Pxml).WriteInfo("functionName : " + tfunctionName);
                        new CustomLogger.Logger(pM2Pxml).WriteInfo("Search UserID : " + tStrUserID);

                        //取公司別值
                        tResponse = CallAPI(pM2Pxml, tfunctionName, tparam, out tErrorMsg);
                        if (!string.IsNullOrEmpty(tErrorMsg)) return ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);

                        tResponse = "{\"CompanyItems\" : " + tResponse + "}";                                           //為符合class格式自己加上去的
                        GetCompany CompanyClass = tSerObject.Deserialize<GetCompany>(tResponse);                        //轉class方便取用
                        DataTable tCompanyTable = CompanyClass.GetCompanyTable(EmployDetail.company);                   //取公司別table
                        DataRow[] tCompanyRow = tCompanyTable.Select("company = '" + EmployDetail.company + "'");       //查詢指定company
                        if (tCompanyRow.Count() > 0) tCompany = tCompanyRow[0]["company"] + "-" + tCompanyRow[0]["name"];
                        #endregion

                        #endregion

                        #region 給值
                        //設定加班單資料
                        DocDataClass = new Docdata();
                        DocDataClass.head = new HeadData(tCompany, tDate, tUser, tDetp, tStrformType, fStrormTypeDesc);
                        string DocDataStr = tSerObject.Serialize(DocDataClass);//Class轉成JSON

                        //給定畫面欄位值
                        tRCPcompany.Value = tCompany;       //公司別
                        tRCPdate.Value = tDate;             //申請日
                        tRCPuserid.Value = tUser;           //申請人
                        tRCPdeptno.Value = tDetp;           //申請部門
                        tRCPDocData.Value = DocDataStr;     //加班單暫存資料
                        tRCPBodyInfo.Table = new DataTable();//加班單單身
                        #endregion

                        //處理回傳
                        tP2MObject.AddRCPControls(tRCPformType, tRCPcompany, tRCPdate, tRCPuserid, tRCPdeptno, tRCPBodyInfo, tRCPDocData);
                    }
                    else
                    {
                        tP2MObject.Message = "請假單建立失敗，說明 :\r\n" + tCreateDocResultClass.msg;
                        tP2MObject.Result = "false";
                    }
                }
                #endregion

            }
            catch (Exception err)
            {
                ReturnErrorMsg(pM2Pxml, ref tP2MObject, "Get_HANBELL01_CreateDoc Error : " + err.Message.ToString());
            }

            tP2Mxml = tP2MObject.ToDucument().ToString();
            return tP2Mxml;
        }
        #endregion

        //-----------------------------------------------------------------------------------------------------------------------------------
        #region 加班明細 頁面初使化
        public string Get_HANBELL01_01_BasicSetting(XDocument pM2Pxml)
        {
            string tP2Mxml = string.Empty;//回傳值
            ProductToMCloudBuilder tP2MObject = new ProductToMCloudBuilder(null, null, null); //參數:要顯示的訊息、結果、Title//建構p2m

            try
            {
                #region 設定參數                        
                JavaScriptSerializer tSerObject = new JavaScriptSerializer();  //設定JSON轉換物件
                Dictionary<string, string> tparam = new Dictionary<string, string>();//叫用API的參數
                string tfunctionName = string.Empty;    //叫用API的方法名稱(自定義)
                string tErrorMsg = string.Empty;        //檢查是否錯誤
                #endregion

                #region 設定控件
                RCPControl tRCPAdd = new RCPControl("Add", "true", "true",  null);          //新增按鈕
                RCPControl tRCPEdit = new RCPControl("Edit", "true", "false",  null);       //修改
                RCPControl tRCfood = new RCPControl("food", "true", null, null);            //供餐
                RCPControl tRCPdeptno = new RCPControl("deptno", "true", null, null);       //加班部門
                RCPControl tRCPids = new RCPControl("ids","true", "true",  null);           //加班人員(多人/新增)
                RCPControl tRCPid = new RCPControl("id","false", "false",  null);           //加班人員(單人/修改)
                RCPControl tRCPdate = new RCPControl("date", "true", null, null);           //加班日期
                RCPControl tRCPstarttime = new RCPControl("starttime", "true", null, null); //加班起
                RCPControl tRCPendtime = new RCPControl("endtime", "true", null, null);     //加班迄
                RCPControl tRCPhour = new RCPControl("hour", "true", null, null);           //加班時數
                RCPControl tRCPnote = new RCPControl("note", "true", null, null);           //加班說明
                RCPControl tRCPDocData = new RCPControl("DocData", null, null, null);       //申請單資料(隱)
                RCPControl tRCPitem = new RCPControl("item", null, null, null);             //判斷是新增或是修改(隱)
                #endregion

                #region 取得畫面資料
                string tStrLanguage = DataTransferManager.GetDataValue(pM2Pxml, "Language");        //語系
                string tStrUserID = DataTransferManager.GetUserMappingValue(pM2Pxml, "HANBELL");        //登入者帳號
                string tStrDocData = DataTransferManager.GetControlsValue(pM2Pxml, "DocData");          //加班單暫存資料
                string tStritem = DataTransferManager.GetControlsValue(pM2Pxml, "item");                //判斷是新增或是修改
                #endregion

                #region 處裡畫面資料
                //判斷新增或是修改
                if (string.IsNullOrEmpty(tStritem))//新增
                {
                    #region 取得預設資料
                    //設定參數
                    tfunctionName = "GetEmployeeDetail";    //取得員工明細資料
                    tparam.Clear();
                    tparam.Add("UserID", tStrUserID);       //查詢對象
                    new CustomLogger.Logger(pM2Pxml).WriteInfo("functionName : " + tfunctionName);
                    new CustomLogger.Logger(pM2Pxml).WriteInfo("Search UserID : " + tStrUserID);

                    //叫用服務
                    string tResponse = CallAPI(pM2Pxml, tfunctionName, tparam, out tErrorMsg);
                    if (!string.IsNullOrEmpty(tErrorMsg)) return ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);

                    //回傳的JSON轉Class
                    GetEmployeeDetail EmployDetail = tSerObject.Deserialize<GetEmployeeDetail>(tResponse);  //員工明細

                    //取值
                    string tDetp = string.Format("{0}-{1}", EmployDetail.deptno, EmployDetail.deptname);                //部門代號-部門名稱
                    string tDate = DateTime.Now.ToString("yyyy/MM/dd"); //加班日期
                    string tTime = DateTime.Now.ToString("HH:mm");//加班起訖
                    #endregion

                    #region 給預設值
                    tRCfood.Value = "0";                //供餐
                    tRCPdeptno.Value = tDetp;           //加班部門
                    tRCPids.Value = "";                 //加班人員(多人/新增)
                    tRCPid.Value = "";                  //加班人員(單人/修改)
                    tRCPdate.Value = tDate;             //加班日期
                    tRCPstarttime.Value = tTime;        //加班起
                    tRCPendtime.Value = tTime;          //加班迄
                    tRCPhour.Value = "";                //加班時數
                    tRCPnote.Value = "";                //加班說明
                    tRCPDocData.Value = tStrDocData;    //申請單資料(隱)
                    tRCPitem.Value = tStritem;          //判斷是新增或是修改(隱)                    
                    #endregion

                    #region 設定物件屬性
                    tRCPAdd.Visible = "true";           //新增按鈕
                    tRCPEdit.Visible = "false";         //修改按鈕
                    tRCPdeptno.Enable = "true";         //加班部門
                    tRCPids.Visible = "true";           //加班人員(多人/新增)
                    tRCPid.Visible = "false";           //加班人員(單人/修改)
                    #endregion                    
                }
                else //修改
                {
                    #region 取得單身資料
                    //先將暫存資料轉成Class方便取用
                    Docdata DocDataClass = tSerObject.Deserialize<Docdata>(tStrDocData);//加班單暫存資料

                    //取值
                    string tdinner = string.Empty;      //晚餐
                    string tlunch = string.Empty;       //午餐
                    string tdeptno = string.Empty;      //加班部門
                    string tid = string.Empty;          //加班人員
                    string tdate = string.Empty;        //加班日期
                    string tstarttime = string.Empty;   //加班起
                    string tendtime = string.Empty;     //加班迄
                    string tworktime = string.Empty;    //加班時數
                    string tnote = string.Empty;        //加班說明    

                    foreach (var bodyitem in DocDataClass.body)
                    {
                        if (tStritem.Equals(bodyitem.item))
                        {
                            tdinner = bodyitem.dinner;
                            tlunch = bodyitem.lunch;
                            tdeptno = bodyitem.deptno;
                            tid = bodyitem.id;
                            tdate = bodyitem.date;
                            tstarttime = bodyitem.starttime;
                            tendtime = bodyitem.endtime;
                            tworktime = bodyitem.worktime;
                            tnote = bodyitem.note;
                            break;
                        }
                    }
                    #endregion

                    #region 給值
                    tRCfood.Value = //供餐
                        ("N".Equals(tlunch) && "N".Equals(tdinner)) ? "0" :         //不供餐
                        ("Y".Equals(tlunch) && "Y".Equals(tdinner)) ? "3" :         //供午、晚餐
                        ("Y".Equals(tlunch) && "N".Equals(tdinner)) ? "1" : "2";    //午餐或晚餐
                    tRCPdeptno.Value = tdeptno;         //加班部門
                    tRCPids.Value = "";                 //加班人員(多人/新增)
                    tRCPid.Value = tid;                 //加班人員(單人/修改)
                    tRCPdate.Value = tdate;             //加班日期
                    tRCPstarttime.Value = tstarttime;   //加班起
                    tRCPendtime.Value = tendtime;       //加班迄
                    tRCPhour.Value = tworktime;         //加班時數
                    tRCPnote.Value = tnote;             //加班說明
                    tRCPDocData.Value = tStrDocData;    //申請單資料(隱)
                    tRCPitem.Value = tStritem;          //判斷是新增或是修改(隱)                    
                    #endregion

                    #region 設定物件屬性
                    tRCPAdd.Visible = "false";          //新增按鈕
                    tRCPEdit.Visible = "true";          //修改按鈕
                    tRCPdeptno.Enable = "false";        //加班部門
                    tRCPids.Visible = "false";          //加班人員(多人/新增)
                    tRCPid.Visible = "true";            //加班人員(單人/修改)
                    #endregion
                }
                #endregion

                #region 設定物件屬性
                //設定多語系
                if ("Lang01".Equals(tStrLanguage))
                {
                    //繁體
                    tRCPAdd.Title = "新增";
                    tRCPEdit.Title = "修改";
                    tRCfood.Title = "供餐";
                    tRCPdeptno.Title = "加班部門";
                    tRCPids.Title = "加班人員";
                    tRCPid.Title = "加班人員";
                    tRCPdate.Title = "加班日期";
                    tRCPstarttime.Title = "加班起";
                    tRCPendtime.Title = "加班迄";
                    tRCPhour.Title = "加班時數";
                    tRCPnote.Title = "加班說明";
                }
                else
                {
                    //簡體
                    tRCPAdd.Text = "新增";
                    tRCPEdit.Text = "修改";
                    tRCfood.Title = "供餐";
                    tRCPdeptno.Title = "加班部门";
                    tRCPids.Title = "加班人员";
                    tRCPid.Title = "加班人员";
                    tRCPdate.Title = "加班日期";
                    tRCPstarttime.Title = "加班起";
                    tRCPendtime.Title = "加班迄";
                    tRCPhour.Title = "加班时数";
                    tRCPnote.Title = "加班说明";
                }
                #endregion

                //處理回傳
                tP2MObject.AddTimeout(300);
                tP2MObject.AddRCPControls(tRCPAdd, tRCPEdit, tRCfood, tRCPdeptno, tRCPids, tRCPid, tRCPdate, 
                                          tRCPstarttime, tRCPendtime, tRCPhour, tRCPnote, tRCPDocData, tRCPitem);
            }
            catch (Exception err)
            {
                ReturnErrorMsg(pM2Pxml, ref tP2MObject, "Get_HANBELL01_01_BasicSetting Error : " + err.Message.ToString());
            }

            tP2Mxml = tP2MObject.ToDucument().ToString();
            return tP2Mxml;
        }
        #endregion

        #region 部門開窗
        public string Get_HANBELL01_01_deptno_OP(XDocument pM2Pxml)
        {
            string tP2Mxml = string.Empty;//回傳值
            ProductToMCloudBuilder tP2MObject = new ProductToMCloudBuilder(null, null, null); //參數:要顯示的訊息、結果、Title//建構p2m

            try
            {
                #region 設定參數                        
                JavaScriptSerializer tSerObject = new JavaScriptSerializer();  //設定JSON轉換物件
                Dictionary<string, string> tparam = new Dictionary<string, string>();//叫用API的參數
                string tfunctionName = string.Empty;    //叫用API的方法名稱(自定義)
                string tErrorMsg = string.Empty;        //檢查是否錯誤
                #endregion

                #region 設定控件
                RDControl tRDdeptno = new RDControl(new DataTable());  //部門清單
                #endregion

                #region 取得畫面資料
                string tStrUserID = DataTransferManager.GetUserMappingValue(pM2Pxml, "HANBELL");    //登入者帳號
                string tSearch = DataTransferManager.GetControlsValue(pM2Pxml, "SearchCondition");  //單身搜尋條件
                string tStrLanguage = DataTransferManager.GetDataValue(pM2Pxml, "Language");        //語系
                #endregion

                #region 處理畫面資料
                //因部門資料太多, 改成一定要下搜尋條件才顯示資料
                if (!string.IsNullOrEmpty(tSearch))
                {
                    #region 取得部門清單資料
                    //設定參數
                    tfunctionName = "GetDepartment";            //查詢部門
                    tparam.Clear();
                    //tparam.Add("deptId", tSearch);              //部門id
                    tparam.Add("deptName", tSearch);            //部門名稱
                    new CustomLogger.Logger(pM2Pxml).WriteInfo("functionName : " + tfunctionName);

                    //叫用服務
                    string tResponse = CallAPI(pM2Pxml, tfunctionName, tparam, out tErrorMsg);
                    if (!string.IsNullOrEmpty(tErrorMsg)) return ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);
                    tResponse = "{\"DepartmentItems\" : " + tResponse + "}";                            //為符合class格式自己加上去的
                    GetDepartment DeparmentClass = tSerObject.Deserialize<GetDepartment>(tResponse);    //轉class方便取用
                    #endregion

                    #region 給值
                    //取得部門Table
                    DataTable tDeparmentTable = DeparmentClass.GetDeparmentTable(tSearch);

                    //轉換成符合格式的table
                    if (tDeparmentTable.Columns.Contains("organizationUnitName")) { tDeparmentTable.Columns["organizationUnitName"].ColumnName = "deptnoName"; }
                    tDeparmentTable.Columns.Add("deptno");
                    tDeparmentTable.Columns.Add("deptnoC");
                    for (int i = 0; i < tDeparmentTable.Rows.Count; i++) {
                        tDeparmentTable.Rows[i]["deptno"] = tDeparmentTable.Rows[i]["id"] + "-" + tDeparmentTable.Rows[i]["deptnoName"];    //內存
                        tDeparmentTable.Rows[i]["deptnoC"] = tDeparmentTable.Rows[i]["id"] + "-" + tDeparmentTable.Rows[i]["deptnoName"];   //外顯
                    }

                    //設定公司別清單資料(含分頁)
                    SetRDControlpage(pM2Pxml, tDeparmentTable, ref tRDdeptno);
                    #endregion
                }
                else
                {
                    tP2MObject.Message = "請先輸入部門名稱條件";
                    tP2MObject.Result = "true";
                }
                #endregion

                #region 設定物件屬性
                //公司別清單屬性
                tRDdeptno.AddCustomTag("DisplayColumn", "id§deptnoName§deptno");
                tRDdeptno.AddCustomTag("ColumnsName", "代號§名稱§部門");
                tRDdeptno.AddCustomTag("StructureStyle", "id:R1:C1:1§deptnoName:R1:C2:1");

                //設定多語系
                if ("Lang01".Equals(tStrLanguage))
                {
                    //繁體
                    tRDdeptno.AddCustomTag("ColumnsName", "代號§名稱§部門");
                }
                else
                {
                    //簡體
                    tRDdeptno.AddCustomTag("ColumnsName", "代号§名称§部门");
                }
                #endregion

                //處理回傳
                tP2MObject.AddRDControl(tRDdeptno);

            }
            catch (Exception err)
            {
                ReturnErrorMsg(pM2Pxml, ref tP2MObject, "Get_HANBELL01_01_deptno_OP Error : " + err.Message.ToString());
            }

            tP2Mxml = tP2MObject.ToDucument().ToString();
            return tP2Mxml;
        }
        #endregion

        #region 部門異動
        public string Get_HANBELL01_01_deptno_OnBlur(XDocument pM2Pxml)
        {
            string tP2Mxml = string.Empty;//回傳值
            ProductToMCloudBuilder tP2MObject = new ProductToMCloudBuilder(null, null, null); //參數:要顯示的訊息、結果、Title//建構p2m

            try
            {
                #region 設定控件
                RCPControl tRCPids = new RCPControl("ids", null, null, ""); //加班人員 //部門影響人員所以要清空
                #endregion

                //處理回傳
                tP2MObject.AddRCPControls(tRCPids);

            }
            catch (Exception err)
            {
                ReturnErrorMsg(pM2Pxml, ref tP2MObject, "Get_HANBELL01_01_deptno_OnBlur Error : " + err.Message.ToString());
            }

            tP2Mxml = tP2MObject.ToDucument().ToString();
            return tP2Mxml;
        }
        #endregion

        #region 人員開窗
        public string Get_HANBELL01_01_ids_OP(XDocument pM2Pxml)
        {
            string tP2Mxml = string.Empty;//回傳值
            ProductToMCloudBuilder tP2MObject = new ProductToMCloudBuilder(null, null, null); //參數:要顯示的訊息、結果、Title//建構p2m

            try
            {
                #region 設定參數                        
                JavaScriptSerializer tSerObject = new JavaScriptSerializer();  //設定JSON轉換物件
                Dictionary<string, string> tparam = new Dictionary<string, string>();//叫用API的參數
                string tfunctionName = string.Empty;    //叫用API的方法名稱(自定義)
                string tErrorMsg = string.Empty;        //檢查是否錯誤
                #endregion

                #region 設定控件
                RDControl tRDids = new RDControl(new DataTable());  //人員清單
                #endregion

                #region 取得畫面資料
                string tStrUserID = DataTransferManager.GetUserMappingValue(pM2Pxml, "HANBELL");    //登入者帳號
                string tStrdeptno = DataTransferManager.GetControlsValue(pM2Pxml, "deptno");        //查詢的部門
                string tStrids = DataTransferManager.GetControlsValue(pM2Pxml, "ids");              //已勾選人員
                string tSearch = DataTransferManager.GetControlsValue(pM2Pxml, "SearchCondition");  //單身搜尋條件
                string tStrLanguage = DataTransferManager.GetDataValue(pM2Pxml, "Language");        //語系
                #endregion

                #region 處裡畫面資料
                #region 取得資料
                //設定參數
                tfunctionName = "GetDepartmentEmployee";    //取得員工明細資料
                tparam.Clear();
                tparam.Add("deptId", tStrdeptno.Split('-')[0]);       //查詢部門
                new CustomLogger.Logger(pM2Pxml).WriteInfo("functionName : " + tfunctionName);
                new CustomLogger.Logger(pM2Pxml).WriteInfo("Search deptId : " + tStrdeptno.Split('-')[0]);

                //叫用服務
                string tResponse = CallAPI(pM2Pxml, tfunctionName, tparam, out tErrorMsg);
                if (!string.IsNullOrEmpty(tErrorMsg)) return ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);

                tResponse = "{\"DepartmentEmployee\" : " + tResponse + "}";                                         //為符合class格式自己加上去的
                GetDepartmentEmployee EmployeeClass = tSerObject.Deserialize<GetDepartmentEmployee>(tResponse);     //轉class方便取用
                #endregion

                #region 給值
                //取得部門人員Table
                DataTable tEmployeeTable = EmployeeClass.GetDeparmentTable(tSearch);
                tEmployeeTable.Columns.Add("ids");
                for (int i = 0; i < tEmployeeTable.Rows.Count; i++)
                {
                    tEmployeeTable.Rows[i]["ids"] = tEmployeeTable.Rows[i]["id"] + "-" + tEmployeeTable.Rows[i]["userName"];    //KEYValue, 主要讓UIFLOW可以傳到前一頁
                }

                //設定部門人員清單資料(含分頁)
                SetRDControlpage(pM2Pxml, tEmployeeTable, ref tRDids);
                
                #endregion
                #endregion

                #region 設定物件屬性
                //單身清單屬性
                tRDids.AddCustomTag("DisplayColumn", "id§userName");
                tRDids.AddCustomTag("ColumnsName", "工號§姓名");
                tRDids.AddCustomTag("StructureStyle", "id:R1:C1:1§userName:R1:C2:1");

                //設定多語系
                if ("Lang01".Equals(tStrLanguage))
                {
                    //繁體
                    tRDids.AddCustomTag("ColumnsName", "工號§姓名");
                }
                else
                {
                    //簡體
                    tRDids.AddCustomTag("ColumnsName", "工号§姓名");
                }
                #endregion

                //處理回傳
                tP2MObject.AddRDControl(tRDids);

            }
            catch (Exception err)
            {
                ReturnErrorMsg(pM2Pxml, ref tP2MObject, "Get_HANBELL01_01_ids_OP Error : " + err.Message.ToString());
            }

            tP2Mxml = tP2MObject.ToDucument().ToString();
            return tP2Mxml;
        }
        #endregion

        #region 新增/修改單身
        public string Get_HANBELL01_01_Add(XDocument pM2Pxml)
        {
            string tP2Mxml = string.Empty;//回傳值
            ProductToMCloudBuilder tP2MObject = new ProductToMCloudBuilder(null, null, null); //參數:要顯示的訊息、結果、Title//建構p2m

            try
            {
                #region 設定參數                        
                JavaScriptSerializer tSerObject = new JavaScriptSerializer();  //設定JSON轉換物件
                Dictionary<string, string> tparam = new Dictionary<string, string>();//叫用API的參數
                string tfunctionName = string.Empty;    //叫用API的方法名稱(自定義)
                string tErrorMsg = string.Empty;        //檢查是否錯誤
                #endregion

                #region 設定控件
                RCPControl tRCfood = new RCPControl("food", "true", null, null);            //供餐
                RCPControl tRCPdeptno = new RCPControl("deptno", "true", null, null);       //加班部門
                RCPControl tRCPids = new RCPControl("ids", "true", null, null);             //加班人員(多人/新增)
                RCPControl tRCPdate = new RCPControl("date", "true", null, null);           //加班日期
                RCPControl tRCPstarttime = new RCPControl("starttime", "true", null, null); //加班起
                RCPControl tRCPendtime = new RCPControl("endtime", "true", null, null);     //加班迄
                RCPControl tRCPhour = new RCPControl("hour", "true", null, null);           //加班時數
                RCPControl tRCPnote = new RCPControl("note", "true", null, null);           //加班說明
                RCPControl tRCPDocData = new RCPControl("DocData", null, null, null);       //申請單資料(隱)
                #endregion

                #region 取得畫面資料
                string tServiceName = DataTransferManager.GetDataValue(pM2Pxml, "ServiceName");         //service名稱

                string tStrUserID = DataTransferManager.GetUserMappingValue(pM2Pxml, "HANBELL");        //登入者帳號
                string tStrDocData = DataTransferManager.GetControlsValue(pM2Pxml, "DocData");          //加班單暫存資料
                string tStrfood = DataTransferManager.GetControlsValue(pM2Pxml, "food");                //供餐
                string tStrdeptno = DataTransferManager.GetControlsValue(pM2Pxml, "deptno");            //加班部門
                string tStrdate = DataTransferManager.GetControlsValue(pM2Pxml, "date");                //加班日期
                string tStrids = DataTransferManager.GetControlsValue(pM2Pxml, "ids");                  //新增人員畫面回傳的勾選資料
                string tStrstarttime = DataTransferManager.GetControlsValue(pM2Pxml, "starttime");      //加班起
                string tStrendtime = DataTransferManager.GetControlsValue(pM2Pxml, "endtime");          //加班迄
                string tStrhour = DataTransferManager.GetControlsValue(pM2Pxml, "hour");                //加班時數
                string tStrnote = DataTransferManager.GetControlsValue(pM2Pxml, "note");                //加班說明
                string tStritem = DataTransferManager.GetControlsValue(pM2Pxml, "item");                //判斷是新增或是修改
                string tStrLanguage = DataTransferManager.GetDataValue(pM2Pxml, "Language");            //語系

                //先將暫存資料轉成Class方便取用
                Docdata DocDataClass = tSerObject.Deserialize<Docdata>(tStrDocData);//加班單暫存資料
                #endregion

                #region 處裡畫面資料
                #region 驗證輸入資料是否有問題
                //必填
                if (string.IsNullOrEmpty(tStrdeptno))
                {
                    if ("Lang01".Equals(tStrLanguage))
                    {
                        //繁體
                        tErrorMsg += "加班部門必填\r\n";
                    }
                    else
                    {
                        //簡體
                        tErrorMsg += "加班部门必填\r\n";
                    }
                }
                    
                if (string.IsNullOrEmpty(tStrdate))
                {
                    if ("Lang01".Equals(tStrLanguage))
                    {
                        //繁體
                        tErrorMsg += "加班日期必填\r\n";
                    }
                    else
                    {
                        //簡體
                        tErrorMsg += "加班日期必填\r\n";
                    }
                }

                if ("Add".Equals(tServiceName) && string.IsNullOrEmpty(tStrids))
                {
                    if ("Lang01".Equals(tStrLanguage))
                    {
                        //繁體
                        tErrorMsg += "加班人員必填\r\n";
                    }
                    else
                    {
                        //簡體
                        tErrorMsg += "加班人员必填\r\n";
                    }
                }

                if (string.IsNullOrEmpty(tStrhour))
                {
                    if ("Lang01".Equals(tStrLanguage))
                    {
                        //繁體
                        tErrorMsg += "加班時數必填\r\n";
                    }
                    else
                    {
                        //簡體
                        tErrorMsg += "加班时数必填\r\n";
                    }
                }

                if (string.IsNullOrEmpty(tStrnote))
                {
                    if ("Lang01".Equals(tStrLanguage))
                    {
                        //繁體
                        tErrorMsg += "加班說明必填\r\n";
                    }
                    else
                    {
                        //簡體
                        tErrorMsg += "加班说明必填\r\n";
                    }
                }
                if (!string.IsNullOrEmpty(tErrorMsg)) return ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);

                //資料驗證
                //日期
                DateTime thead_Date = Convert.ToDateTime(DocDataClass.head.date);                   //申請日期
                DateTime tbody_Date = Convert.ToDateTime(tStrdate.Insert(4,"/").Insert(7,"/"));     //加班日期
                TimeSpan ts = tbody_Date - thead_Date;                                              //差異天數
                if(ts.TotalDays < -5)
                {
                    if ("Lang01".Equals(tStrLanguage))
                    {
                        //繁體
                        tErrorMsg += "「加班日期」需於「申請日」前五天之後\r\n";
                    }
                    else
                    {
                        //簡體
                        tErrorMsg += "「加班日期」需于「申请日」前五天之后\r\n";
                    }
                }

                //加班起迄
                if(!(int.Parse(tStrendtime) >= int.Parse(tStrstarttime)))
                {
                    if ("Lang01".Equals(tStrLanguage))
                    {
                        //繁體
                        tErrorMsg += "「加班起」需於「加班迄」之前\r\n";
                    }
                    else
                    {
                        //簡體
                        tErrorMsg += "「加班起」需于「加班迄」之前\r\n";
                    }
                }

                //加班時數
                if(!(24 >= double.Parse(tStrhour)  && double.Parse(tStrhour) > 0))
                {
                    if ("Lang01".Equals(tStrLanguage))
                    {
                        //繁體
                        tErrorMsg += "「加班時數」需「小於等於24且大於等於0」\r\n";
                    }
                    else
                    {
                        //簡體
                        tErrorMsg += "「加班时数」需「小於等於24且大於等於0」\r\n";
                    }
                }

                if (!string.IsNullOrEmpty(tErrorMsg)) return ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);
                #endregion

                switch (tServiceName)
                {
                    case "Add":             //新增單身
                        #region 新增單身資料
                        //依人員逐一新增
                        foreach (var employee in tStrids.Split('§'))
                        {
                            string tKeyField = employee + tStrdate; //key值

                            //檢查是否存在
                            foreach (var bodyitem in DocDataClass.body)
                            {
                                if(tKeyField.Equals(bodyitem.item))
                                {
                                    tErrorMsg += employee + "\r\n";
                                }
                            }

                            //檢查不存在則新增
                            if (string.IsNullOrEmpty(tErrorMsg))
                            {
                                BodyData tbody = new BodyData();
                                tbody.item = tKeyField;   //key
                                tbody.lunch = "1".Equals(tStrfood) || "3".Equals(tStrfood) ? "Y" : "N";     //午餐
                                tbody.dinner = "2".Equals(tStrfood) || "3".Equals(tStrfood) ? "Y" : "N";    //晚餐
                                tbody.deptno = tStrdeptno;          //加班部門
                                tbody.date = tStrdate;              //加班日期
                                tbody.id = employee;                //加班人員
                                tbody.starttime = tStrstarttime;    //加班起
                                tbody.endtime = tStrendtime;        //加班迄
                                tbody.worktime = tStrhour;          //加班時數
                                tbody.note = tStrnote;              //加班說明

                                DocDataClass.body.Add(tbody);
                            }
                        }

                        //有重的資料就回傳錯誤
                        if (!string.IsNullOrEmpty(tErrorMsg))
                        {
                            tErrorMsg += tStrdate.Insert(4, "/").Insert(7, "/") + "加班資料已存在";
                            return ReturnErrorMsg(pM2Pxml, ref tP2MObject, tErrorMsg);
                        }
                        #endregion
                        break;
                    case "Edit":            //修改明細
                        #region 修改單身資料
                        foreach (var bodyitem in DocDataClass.body)
                        {
                            if (tStritem.Equals(bodyitem.item))
                            {
                                bodyitem.lunch = "1".Equals(tStrfood) || "3".Equals(tStrfood) ? "Y" : "N";     //午餐
                                bodyitem.dinner = "2".Equals(tStrfood) || "3".Equals(tStrfood) ? "Y" : "N";    //晚餐
                                bodyitem.date = tStrdate;              //加班日期
                                bodyitem.starttime = tStrstarttime;    //加班起
                                bodyitem.endtime = tStrendtime;        //加班迄
                                bodyitem.worktime = tStrhour;          //加班時數
                                bodyitem.note = tStrnote;              //加班說明
                                break;
                            }
                        }
                        #endregion
                        break;
                }

                #region 更新畫面欄位資料
                tRCPDocData.Value = tSerObject.Serialize(DocDataClass);     //加班單暫存
                #endregion

                #endregion

                //處理回傳
                tP2MObject.AddStatus("callwork&HANBELL01_01-HANBELL01+加班明細返回&None"); //返回申請單
                tP2MObject.AddRCPControls(tRCfood, tRCPdeptno, tRCPids, tRCPdate, tRCPstarttime,
                                          tRCPendtime, tRCPhour, tRCPnote, tRCPDocData);
            }
            catch (Exception err)
            {
                ReturnErrorMsg(pM2Pxml, ref tP2MObject, "Get_HANBELL01_01_Add Error : " + err.Message.ToString());
            }

            tP2Mxml = tP2MObject.ToDucument().ToString();
            return tP2Mxml;
        }
        #endregion
                
        //====================================================================================================================================

        #region 分頁設定 (RCPControl)
        /// <summary>
        /// 分頁設定
        /// </summary>
        /// <param name="pM2PXml">M2P</param>
        /// <param name="pDT">GridView內容</param>
        /// <param name="pRCP">GridView控件</param>
        public void SetRCPControlpage(XDocument pM2PXml, DataTable pDT, ref RCPControl pRCP)
        {
            #region 處理分頁
            int tPageNO = 1; //目前頁數(預設1)
            if (DataTransferManager.GetControlsValue(pM2PXml, "PageNo") != string.Empty)
            {
                tPageNO = Convert.ToInt16(DataTransferManager.GetControlsValue(pM2PXml, "PageNo"));  //檢查頁數
            }
            double tQueryCount = 20;//顯示筆數

            double tDdouble = pDT.Rows.Count / tQueryCount;
            int tTotalPage = Convert.ToInt32(Math.Ceiling(tDdouble));//總頁數

            #endregion

            //設定屬性
            pRCP.AddCustomTag("TotalPage", tTotalPage.ToString());      //總頁數
            pRCP.AddCustomTag("PageNo", tPageNO.ToString());            //目前頁數
            pRCP.Table = DataTransferManager.MakeDataTablePaged(pDT, tPageNO, Convert.ToInt32(tQueryCount));    //把結果給GridView
        }
        #endregion

        #region 分頁設定 (RDControl)
        /// <summary>
        /// 分頁設定
        /// </summary>
        /// <param name="pM2PXml">M2P</param>
        /// <param name="pDT">GridView內容</param>
        /// <param name="pRCP">GridView控件</param>
        public void SetRDControlpage(XDocument pM2PXml, DataTable pDT, ref RDControl pRCP)
        {
            #region 處理分頁
            int tPageNO = 1; //目前頁數(預設1)
            if (DataTransferManager.GetControlsValue(pM2PXml, "PageNo") != string.Empty)
            {
                tPageNO = Convert.ToInt16(DataTransferManager.GetControlsValue(pM2PXml, "PageNo"));  //檢查頁數
            }
            double tQueryCount = 20;//顯示筆數

            double tDdouble = pDT.Rows.Count / tQueryCount;
            int tTotalPage = Convert.ToInt32(Math.Ceiling(tDdouble));//總頁數

            #endregion

            //設定屬性
            pRCP.AddCustomTag("TotalPage", tTotalPage.ToString());      //總頁數
            pRCP.AddCustomTag("PageNo", tPageNO.ToString());            //目前頁數
            pRCP.Table = DataTransferManager.MakeDataTablePaged(pDT, tPageNO, Convert.ToInt32(tQueryCount));    //把結果給GridView
        }
        #endregion

        #region 處理報錯
        /// <summary>
        /// 處理報錯
        /// </summary>
        /// <param name="pM2Pxml">M2P</param>
        /// <param name="pP2MObject">P2M</param>
        /// <param name="pErrorMsg">錯誤訊息</param>
        /// <returns></returns>
        public string ReturnErrorMsg(XDocument pM2Pxml, ref ProductToMCloudBuilder pP2MObject, string pErrorMsg)
        {
            pP2MObject.Message = pErrorMsg;
            pP2MObject.Result = "false";
            new CustomLogger.Logger(pM2Pxml).WriteError(pErrorMsg);

            return pP2MObject.ToDucument().ToString();
        }
        #endregion

        #region CallAPI
        /// <summary>
        /// 叫用API
        /// </summary>
        /// <param name="pM2Pxml">M2P</param>
        /// <param name="pFunctionName">服務名稱(請看 取得Request)</param>
        /// <param name="pParam">取得Request的參數</param>
        /// <param name="pErrorMsg">錯誤訊息</param>
        /// <returns></returns>
        public string CallAPI(XDocument pM2Pxml, string pFunctionName, Dictionary<string, string> pParam, out string pErrorMsg)
        {
            pErrorMsg = string.Empty;

            string tRequest = GetRequest(pFunctionName, pParam, out pErrorMsg);
            if (!string.IsNullOrEmpty(pErrorMsg))
            {
                pErrorMsg = string.Format("function : {0}\r\nGet Request Error : {1}", pFunctionName, pErrorMsg);
                return "";
            }
            else
            {
                new CustomLogger.Logger(pM2Pxml).WriteInfo("Request : " + tRequest);
            }

            string tResponse = CallAPI(tRequest, out pErrorMsg);
            if (!string.IsNullOrEmpty(pErrorMsg))
            {
                pErrorMsg = string.Format("function : {0}\r\nGet Response Error : {1}", pFunctionName, pErrorMsg);
                return "";
            }
            else
            {
                new CustomLogger.Logger(pM2Pxml).WriteInfo("Response : " + tResponse);
            }

            return tResponse;
        }
        #endregion

        #region 取得Request
        /// <summary>
        /// 取得Request
        /// </summary>
        /// <param name="functionName">方法名稱</param>
        /// <param name="param">參數</param>
        /// <returns></returns>
        public string GetRequest(string functionName, Dictionary<string, string> param, out string pErrorMsg)
        {
            /*使用範例
                Dictionary<string, string> tparam = new Dictionary<string, string>();
                new CustomLogger.Logger(pM2Pxml).WriteInfo("functionName : GetCompany");
                string trequest = GetRequest("GetCompany", tparam);
                new CustomLogger.Logger(pM2Pxml).WriteInfo("Request : " + trequest);
                string tresponse = CallAPI(trequest);
                new CustomLogger.Logger(pM2Pxml).WriteInfo("Response : " + tresponse);
             */
            pErrorMsg = string.Empty;       //錯誤訊息
            string Result = string.Empty;   //服務位置
            try
            {
                switch (functionName)
                {
                    #region 取得公司別
                    case "GetCompany":
                        //ex:http://i2.hanbell.com.cn:8080/Hanbell-JRS/api/eap/company?appid=1505278334853&token=0ec858293fccfad55575e26b0ce31177
                        //取回:公司代号company、公司简称name
                        Result = "eap/company";
                        break;
                    #endregion

                    #region 取得員工清單資料
                    case "GetEmployee":
                        //ex:http://i2.hanbell.com.cn:8080/Hanbell-JRS/api/efgp/users/f/s/0/20?appid=1505278334853&token=0ec858293fccfad55575e26b0ce31177
                        //返回EFGP中从0开始的20个用户(/0/20)，没有筛选和排序条件（/f/s）
                        //ex:http://i2.hanbell.com.cn:8080/Hanbell-JRS/api/efgp/users/f;userName=天雨/s/0/20?appid=1505278334853&token=0ec858293fccfad55575e26b0ce31177
                        //返回EFGP Users表中,姓名包含“天雨”的20条记录，筛选条件/f;userName=天雨,“f”是个固定值，“;”后面是条件，多个条件用“;”隔开
                        //ex:http://i2.hanbell.com.cn:8080/Hanbell-JRS/api/efgp/users/f;userName=王/s;id=DESC/0/20/size?appid=1505278334853&token=0ec858293fccfad55575e26b0ce31177
                        //返回EFGP Users表中,姓名包含“王”的20条记录，并按id字段倒排序，排序条件/s;id=DESC，“s”是个固定值,“;”后面是条件，多个条件用“;”隔开，如果是顺序就是/s;id=ASC
                        //篩選條件:按代号（id=XXXX）,名称（userName=XXXX）
                        //取回:员工代号id、员工姓名userName

                        param.Add("condition",
                            (param.ContainsKey("userId") ? string.Format(";id={0}", param["userId"]) : "") +            //篩選條件:按代号（id=XXXX）
                            (param.ContainsKey("userName") ? string.Format(";userName={0}", param["userName"]) : "")    //篩選條件:按名称（userName=XXXX）
                        );

                        Result = string.Format("efgp/users/f{0}/s;id=ASC/{1}/{2}",
                                    param.ContainsKey("condition") ? param["condition"] : "",   //搜尋條件 (;id=XX;userName=xx)
                                    param.ContainsKey("StartRow") ? param["StartRow"] : "0",    //搜尋起始筆數
                                    param.ContainsKey("EndRow") ? param["EndRow"] : "1000"      //搜尋結尾筆數
                                 );
                        break;
                    #endregion

                    #region 取得員工明細資料
                    case "GetEmployeeDetail":
                        //ex:http://i2.hanbell.com.cn:8080/Hanbell-JRS/api/efgp/users/C0160?appid=1505278334853&token=0ec858293fccfad55575e26b0ce31177
                        //取回:员工代号id、员工姓名userName、主要部门deptno、部门名称deptname、默认公司company
                        Result = string.Format("efgp/users/{0}",
                                    param.ContainsKey("UserID") ? param["UserID"] : "0001"  //使用者待號
                                );
                        break;
                    #endregion

                    #region 查詢員工部門
                    case "GetEmployeeDepartment":
                        //ex:http://i2.hanbell.com.cn:8080/Hanbell-JRS/api/efgp/functions/f;users.id=C0160/s/0/20?appid=1505278334853&token=0ec858293fccfad55575e26b0ce31177
                        //根据Users.id返回EFGP中此账户所在部门
                        //篩選條件:按代号(users.id=xx)
                        //取回:部门代号organizationUnit.id、部门名称organizationUnit.organizationUnitName

                        param.Add("condition",
                            (param.ContainsKey("userId") ? string.Format(";users.id={0}", param["userId"]) : "")  //篩選條件:按代号(users.id=xx)
                        );

                        Result = string.Format("efgp/functions/f{0}/s/{1}/{2}",
                                    param.ContainsKey("condition") ? param["condition"] : "",   //搜尋條件 (;users.id=xx)
                                    param.ContainsKey("StartRow") ? param["StartRow"] : "0",    //搜尋起始筆數
                                    param.ContainsKey("EndRow") ? param["EndRow"] : "1000"      //搜尋結尾筆數
                                );
                        break;
                    #endregion

                    #region 部門查詢
                    case "GetDepartment":
                        //ex:http://i2.hanbell.com.cn:8080/Hanbell-JRS/api/efgp/organizationunit/f;id=13120/s;id=ASC/0/20?appid=1505278334853&token=0ec858293fccfad55575e26b0ce31177
                        //返回EFGP中的部门信息，/f;id=13120/s;id=ASC，筛选和排序条件，没有条件时用/f/s
                        //篩選條件:按代号（id=XXXX）,名称（organizationUnitName=XXXX）
                        //取回:部门编号id、部门名称organizationUnitName

                        param.Add("condition",
                            (param.ContainsKey("deptId") ? string.Format(";id={0}", param["deptId"]) : "") +                        //篩選條件:按代号(id=XXXX)
                            (param.ContainsKey("deptName") ? string.Format(";organizationUnitName={0}", param["deptName"]) : "")    //篩選條件:按名称（organizationUnitName=XXXX）
                        );

                        Result = string.Format("efgp/organizationunit/f{0}/s;id=ASC/{1}/{2}",
                                    param.ContainsKey("condition") ? param["condition"] : "",   //搜尋條件 (;id=xx;organizationUnitName=xx)
                                    param.ContainsKey("StartRow") ? param["StartRow"] : "0",    //搜尋起始筆數
                                    param.ContainsKey("EndRow") ? param["EndRow"] : "1000"      //搜尋結尾筆數
                                );
                        break;
                    #endregion

                    #region 部門人員查詢
                    case "GetDepartmentEmployee":
                        //ex:http://i2.hanbell.com.cn:8080/Hanbell-JRS/api/efgp/users/functions/organizationunit/f;organizationUnit.id=13120/s/0/20?appid=1505278334853&token=0ec858293fccfad55575e26b0ce31177
                        //根据部门代号返回EFGP中此部门相关人员信息
                        //篩選條件:部门代号(organizationUnit.id=XXXX)
                        //取回:员工代号id、员工姓名userName

                        param.Add("condition",
                            (param.ContainsKey("deptId") ? string.Format(";organizationUnit.id={0}", param["deptId"]) : "")  //篩選條件:部门代号(organizationUnit.id=XXXX)
                        );

                        Result = string.Format("efgp/users/functions/organizationunit/f{0}/s/{1}/{2}",
                                    param.ContainsKey("condition") ? param["condition"] : "",   //搜尋條件 (;users.id=xx)
                                    param.ContainsKey("StartRow") ? param["StartRow"] : "0",    //搜尋起始筆數
                                    param.ContainsKey("EndRow") ? param["EndRow"] : "1000"      //搜尋結尾筆數
                                );
                        break;
                        #endregion

                }

                if (!string.IsNullOrEmpty(Result))
                {
                    //組合request
                    Result = string.Format("{0}{1}?{2}", LoadConfig.GetWebConfig("APIURI"), Result, LoadConfig.GetWebConfig("APIKey"));
                }
                else
                {
                    pErrorMsg = "找不到functionName";
                }
            }
            catch(Exception err)
            {
                pErrorMsg = err.Message.ToString();
            }

            return Result;
        }
        #endregion

        #region 取得Response
        /// <summary>
        /// 叫用API
        /// </summary>
        /// <param name="pRequest">Request</param>
        /// <param name="pErrorMsg">錯誤訊息</param>
        /// <param name="pParam">錯誤訊息</param>
        /// <returns>Rsponse</returns>
        public string CallAPI(string pRequest, out string pErrorMsg, params Dictionary<string, string>[] pParam)
        {
            pErrorMsg = string.Empty;   //錯誤訊息
            string tResponse = string.Empty;

            //Get Response
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(pRequest);   //建立連線
                request.Timeout = 300000;                                               //設定timeout時間

                if (pParam.Count() == 0)
                {
                    //一般取response
                    request.Method = "GET";
                }
                else
                {
                    //生單才有額外參數
                    request.Method = "POST";
                    request.ContentType = "application/json";

                    if (pParam[0].ContainsKey("BodyContent"))
                    {
                        using (var streamWriter = new StreamWriter(request.GetRequestStream()))
                        {
                            streamWriter.Write(pParam[0]["BodyContent"]);
                        }
                    }
                    else
                    {
                        pErrorMsg = "叫用服務失敗: 找不到 BodyContent";
                    }
                }

                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)//送出參數取得回傳值
                {
                    if (response.StatusCode == HttpStatusCode.OK)//判斷是否成功
                    {
                        using (Stream reponseStream = response.GetResponseStream())//取得回傳結果
                        {
                            using (StreamReader rd = new StreamReader(reponseStream, Encoding.UTF8))//讀取
                            {
                                string responseString = rd.ReadToEnd();

                                tResponse = responseString;
                            }
                        }
                    }
                    else
                    {
                        pErrorMsg = "叫用服務失敗:" + response.StatusCode.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                pErrorMsg = "叫用服務失敗:" + ex.Message.ToString();
            }

            return tResponse;
        }
        #endregion

    }

    //====================================================================================================================================

    #region 加班單暫存檔處理(Class、JSON轉換處理)
    /*使用範例
     * 
     * 
        JavaScriptSerializer ser = new JavaScriptSerializer();  //設定JSON轉換物件
                                                                //組合head
        HeadData Head = new HeadData("公司別代號", "申請日的欄位", "使用者代號", "部門代號");

        //組合body
        List<BodyData> Body = new List<BodyData>();
        for (int i = 0; i < 3; i++)
        {
            Body.Add(new BodyData("午餐", "晚餐", "加班部門", "加班人員", "加班日期",
                                    "加班起", "加班迄", "加班時數", "備註"));
        }

        Docdata Request_class = new Docdata(Head, Body);  //組合假單Class
        string Request_Str = ser.Serialize(Request_class);//Class轉成JSON

        //JSON轉成Class直接使用
        Docdata Request_Class = ser.Deserialize<Docdata>(Request_Str); //JSON轉Class
        string company = Request_Class.head.company;
        string bodyrowcount = Request_Class.body.Count.ToString();
        string note = Request_Class.body[0].note;
         
    */
    /*JSON樣式
    {
        "head":
        {
            "company":"公司別代號",
            "date":"申請日的欄位,yyyy/MM/dd",
            "id":"申請人,使用者代號",
            "deptno":"申請部門,部門代號"
        }
        ,
        "body":[
        {
            "lunch":"午餐,Y/N",
            "dinner":"晚餐,Y/N",
            "deptno":"加班部門,部門代號",
            "id":"加班人員,員工代號",
            "date":"加班日期,yyyy/MM/dd",
            "starttime":"加班起,hh:mm",
            "endtime":"加班迄,hh:mm",
            "worktime":"加班時數,24>=time>=0.1",
            "note":"備註"
        },
        {
            "lunch":"午餐,Y/N",
            "dinner":"晚餐,Y/N",
            "deptno":"加班部門,部門代號",
            "id":"加班人員,員工代號",
            "date":"加班日期,yyyy/MM/dd",
            "starttime":"加班起,hh:mm",
            "endtime":"加班迄,hh:mm",
            "worktime":"加班時數,24>=time>=0.1",
            "note":"備註"
        }
        ],
        "appid":"1505278334853",
        "token":"0ec858293fccfad55575e26b0ce31177"
    }
    */

    //單頭欄位
    class HeadData
    {
        public string company { get; set; }     //公司別
        public string date { get; set; }        //申請日
        public string id { get; set; }          //申請人
        public string deptno { get; set; }      //申請部門
        public string formType { get; set; }    //加班類別code
        public string formTypeDesc { get; set; }//加班類別name

        public HeadData()
        {
            this.company = "";
            this.date = "";
            this.id = "";
            this.deptno = "";
            this.formType = "0";
            this.formTypeDesc = "平日加班";
        }

        public HeadData(string company, string date, string id, string deptno, string formType, string formTypeDesc)
        {
            this.company = company;
            this.date = date;
            this.id = id;
            this.deptno = deptno;
            this.formType = formType;
            this.formTypeDesc = formTypeDesc;
        }
    }

    //單身欄位
    class BodyData
    {
        public string item { get; set; }        //Key值
        public string lunch { get; set; }       //午餐
        public string dinner { get; set; }      //晚餐
        public string deptno { get; set; }      //加班部門
        public string id { get; set; }          //加班人員
        public string date { get; set; }        //加班日期
        public string starttime { get; set; }   //加班起
        public string endtime { get; set; }     //加班迄
        public string worktime { get; set; }    //加班時數
        public string note { get; set; }        //備註

        public BodyData() { }

        public BodyData(string item, string lunch, string dinner, string deptno, string id, string date,
                        string starttime, string endtime, string worktime, string note)
        {
            this.item = item;
            this.lunch = lunch;
            this.dinner = dinner;
            this.deptno = deptno;
            this.id = id;
            this.date = date;
            this.starttime = starttime;
            this.endtime = endtime;
            this.worktime = worktime;
            this.note = note;
        }

        //取得單身欄位
        public Dictionary<string, string[]> GetBodyDataItem()
        {
            Dictionary<string, string[]> Items = new Dictionary<string, string[]>();
            Items.Add("Columns", new string[] { "item", "lunch", "dinner", "deptno", "id", "date", "starttime", "endtime", "worktime", "note" });
            Items.Add("Values", new string[] { this.item, this.lunch, this.dinner, this.deptno, this.id, this.date, this.starttime, this.endtime, this.worktime, this.note });
                        
            return Items;
        }
    }

    //加班單Class
    class Docdata
    {
        public HeadData head { get; set; }
        public List<BodyData> body { get; set; }

        public Docdata()
        {
            this.head = new HeadData();
            this.body = new List<BodyData>();
        }

        public Docdata(HeadData thead, List<BodyData> body)
        {
            this.head = thead;
            this.body = body;
        }

        //回傳單身table
        public DataTable GetBodyDataTable(string pSearch)
        {
            Dictionary<string, string[]> BodyItems = new Dictionary<string, string[]>();  //取得body有的<欄位/值, 陣列值>

            DataTable BodyTable = new DataTable();
            for (int i = 0; i < this.body.Count; i++)
            {
                bool tInsertData = false;               //是否insert
                BodyItems = this.body[i].GetBodyDataItem(); //取得body資料

                //第一次進來要新增欄位
                if (i == 0) foreach (var col in BodyItems["Columns"]) BodyTable.Columns.Add(col);

                //檢查搜尋
                if (string.IsNullOrEmpty(pSearch))
                {
                    tInsertData = true;
                }
                else
                {
                    //有搜尋字段才需處理搜尋
                    foreach (var value in BodyItems["Values"])
                    {
                        if (value.Contains(pSearch))
                        {
                            tInsertData = true;//找到就跳出
                            break;
                        }
                    }
                }

                //新增資料
                if(tInsertData) BodyTable.Rows.Add(BodyItems["Values"]);
            }

            return BodyTable;
        }
    }
    #endregion

    #region 員工明細
    /*JSON樣式
    {
	    "company": "C",
	    "cost": 0,
	    "deptname": "维护组",
	    "deptno": "13120",
	    "enableSubstitute": 1,
	    "id": "C0160",
	    "identificationType": "DEFAULT",
	    "ldapid": "DongTY",
	    "localeString": "zh_CN",
	    "mailAddress": "dtianyu@hanbell.cn",
	    "mailingFrequencyType": 0,
	    "objectVersion": 15,
	    "oid": "3a6b1714dbbb10048c65af58217aa75e",
	    "password": "s//RxFNKLw5LqwI3D34/DWUdH1w=",
	    "phoneNumber": "65299",
	    "title": {
		    "objectVersion": 1,
		    "occupantOID": "3a6b1714dbbb10048c65af58217aa75e",
		    "oid": "6dfa8be2dd29100482e520de6fdbb7f2",
		    "organizationUnitOID": "3aab5782dbbb10048c65af58217aa75e",
		    "titleDefinition": {
			    "description": "C",
			    "objectVersion": 1,
			    "oid": "6d88eb92dd29100482e520de6fdbb7f2",
			    "organizationOID": "3a8a4bb8dbbb10048c65af58217aa75e",
			    "titleDefinitionName": "C"
		    }
	    },
	    "userName": "董天雨",
	    "workflowServerOID": "00000000000000WorkflowServer0001"
    }
    */

    //titleDefinitionName
    class TITLEDEFINITION
    {
        public string description { get; set; }
        public string objectVersion { get; set; }
        public string oid { get; set; }
        public string organizationOID { get; set; }
        public string titleDefinitionName { get; set; }

        public TITLEDEFINITION() { }
    }

    //TITLE
    class TITLE
    {
        public string objectVersion { get; set; }
        public string occupantOID { get; set; }
        public string oid { get; set; }
        public string organizationUnitOID { get; set; }
        public List<TITLEDEFINITION> titleDefinition { get; set; }

        public TITLE()
        {
            this.titleDefinition = new List<TITLEDEFINITION>();
        }
    }

    //員工明細Class
    class GetEmployeeDetail
    {
        public string company { get; set; }             //默认公司company
        public string cost { get; set; }
        public string deptname { get; set; }            //部门名称deptname
        public string deptno { get; set; }              //主要部门deptno
        public string enableSubstitute { get; set; }
        public string id { get; set; }                  //员工代号id
        public string identificationType { get; set; }
        public string ldapid { get; set; }
        public string localeString { get; set; }
        public string mailAddress { get; set; }
        public string mailingFrequencyType { get; set; }
        public string objectVersion { get; set; }
        public string oid { get; set; }
        public string password { get; set; }
        public string phoneNumber { get; set; }
        public string userName { get; set; }            //员工姓名userName
        public string workflowServerOID { get; set; }
        public List<TITLE> title { get; set; }

        public GetEmployeeDetail()
        {
            this.title = new List<TITLE>();
        }
    }
    #endregion

    #region 公司別清單
    /*JSON樣式
     {//為符合class格式自己加上去的
	    "CompanyItems" ://為符合class格式自己加上去的
        [{
			    "type" : "company",
			    "id" : 1,
			    "cfmdate" : "2017-06-23T22:31:06+08:00",
			    "cfmuser" : "系统管理员",
			    "creator" : "董天雨",
			    "credate" : "2017-05-12T15:03:39+08:00",
			    "optdate" : "2017-06-23T22:30:46+08:00",
			    "optuser" : "系统管理员",
			    "status" : "V",
			    "address" : "",
			    "assetcode" : "S",
			    "boss" : "",
			    "company" : "C",
			    "contacter" : "",
			    "fax" : "57352004",
			    "fullname" : "上海汉钟精机股份有限公司",
			    "name" : "上海汉钟",
			    "remark" : "",
			    "tel" : "57350280"
		    }, {
			    "type" : "company",
			    "id" : 2,
			    "cfmdate" : "2017-06-23T22:31:14+08:00",
			    "cfmuser" : "系统管理员",
			    "creator" : "董天雨",
			    "credate" : "2017-05-12T15:14:28+08:00",
			    "status" : "V",
			    "address" : "",
			    "assetcode" : "G",
			    "boss" : "",
			    "company" : "G",
			    "contacter" : "",
			    "fax" : "",
			    "fullname" : "上海汉钟精机股份有限公司广州分公司",
			    "name" : "汉钟广州",
			    "remark" : "",
			    "tel" : ""
		    }
	    ]
    }//為符合class格式自己加上去的
    */

    //公司別資料
    class CompanyItem
    {
        public string type { get; set; }
        public string id { get; set; }
        public string cfmdate { get; set; }
        public string cfmuser { get; set; }
        public string creator { get; set; }
        public string credate { get; set; }
        public string optdate { get; set; }
        public string optuser { get; set; }
        public string status { get; set; }
        public string address { get; set; }
        public string assetcode { get; set; }
        public string boss { get; set; }
        public string company { get; set; }     //公司代号company、
        public string contacter { get; set; }
        public string fax { get; set; }
        public string fullname { get; set; }
        public string name { get; set; }        //公司简称name
        public string remark { get; set; }
        public string tel { get; set; }

        public CompanyItem() { }

        //取得單身欄位
        public Dictionary<string, string[]> GetCompanyItem()
        {
            Dictionary<string, string[]> Items = new Dictionary<string, string[]>();
            //全部的
            //Items.Add("Columns", new string[] { "type", "id", "cfmdate", "cfmuser", "creator", "credate", "optdate", "optuser", "status", "address",
            //                                    "assetcode", "boss", "company", "contacter", "fax", "fullname", "name", "remark", "tel"});
            //Items.Add("Values", new string[] { this.type, this.id, this.cfmdate, this.cfmuser, this.creator, this.credate, this.optdate, this.optuser, this.status, this.address,
            //                                    this.assetcode, this.boss, this.company, this.contacter, this.fax, this.fullname, this.name, this.remark, this.tel});

            //需要的
            Items.Add("Columns", new string[] { "company", "name"});
            Items.Add("Values", new string[] { this.company, this.name});

            return Items;
        }
    }

    //公司別集合
    class GetCompany
    {
        public List<CompanyItem> CompanyItems { get; set; }

        public GetCompany() {
            this.CompanyItems = new List<CompanyItem>();
        }

        //回傳公司別table
        public DataTable GetCompanyTable(string pSearch)
        {
            Dictionary<string, string[]> Companytems = new Dictionary<string, string[]>();  //取得<欄位/值, 陣列值>

            DataTable CompanyTable = new DataTable();
            for (int i = 0; i < this.CompanyItems.Count; i++)
            {
                bool tInsertData = false;               //是否insert
                Companytems = this.CompanyItems[i].GetCompanyItem(); //取得公司別資料

                //第一次進來要新增欄位
                if (i == 0) foreach (var col in Companytems["Columns"]) CompanyTable.Columns.Add(col);

                //檢查搜尋
                if (string.IsNullOrEmpty(pSearch))
                {
                    tInsertData = true;
                }
                else
                {
                    //有搜尋字段才需處理搜尋
                    foreach (var value in Companytems["Values"])
                    {
                        if (value.Contains(pSearch))
                        {
                            tInsertData = true;//找到就跳出
                            break;
                        }
                    }
                }

                //新增資料
                if (tInsertData) CompanyTable.Rows.Add(Companytems["Values"]);
            }

            return CompanyTable;
        }
    }
    #endregion

    #region 員工部門查詢
    /*JSON樣式
      {//為符合class格式自己加上去的
	    "EmployeeDepartmentItems" : //為符合class格式自己加上去的
        [{
			    "approvalLevelOID" : "3a9aa6d6dbbb10048c65af58217aa75e",
			    "definitionOID" : "3a99fc97dbbb10048c65af58217aa75e",
			    "isMain" : 0,
			    "objectVersion" : 2,
			    "occupantOID" : "3a6b1714dbbb10048c65af58217aa75e",
			    "oid" : "1aee0af6dd0110048c83fb23510cb5dd",
			    "organizationUnit" : {
				    "id" : "13000",
				    "manager" : {
					    "cost" : 0,
					    "enableSubstitute" : 1,
					    "id" : "C0976",
					    "identificationType" : "DEFAULT",
					    "ldapid" : "TangSL",
					    "localeString" : "zh_CN",
					    "mailAddress" : "shunling@hanbell.cn",
					    "mailingFrequencyType" : 0,
					    "objectVersion" : 10,
					    "oid" : "3a6e16b1dbbb10048c65af58217aa75e",
					    "password" : "tyR+kzCgXXHWXKMBvJ81+EP2coI=",
					    "phoneNumber" : "6180",
					    "userName" : "唐舜铃",
					    "workflowServerOID" : "00000000000000WorkflowServer0001"
				    },
				    "objectVersion" : 10,
				    "oid" : "3aab8213dbbb10048c65af58217aa75e",
				    "organizationOID" : "3a8a4bb8dbbb10048c65af58217aa75e",
				    "organizationUnitName" : "资 讯 室",
				    "organizationUnitType" : 0,
				    "superUnitOID" : "3aab026adbbb10048c65af58217aa75e",
				    "validType" : 1
			    },
			    "specifiedManagerOID" : "3a6e16b1dbbb10048c65af58217aa75e",
			    "users" : {
				    "cost" : 0,
				    "enableSubstitute" : 1,
				    "id" : "C0160",
				    "identificationType" : "DEFAULT",
				    "ldapid" : "DongTY",
				    "localeString" : "zh_CN",
				    "mailAddress" : "dtianyu@hanbell.cn",
				    "mailingFrequencyType" : 0,
				    "objectVersion" : 15,
				    "oid" : "3a6b1714dbbb10048c65af58217aa75e",
				    "password" : "s//RxFNKLw5LqwI3D34/DWUdH1w=",
				    "phoneNumber" : "65299",
				    "userName" : "董天雨",
				    "workflowServerOID" : "00000000000000WorkflowServer0001"
			    }
		    }, {
			    "approvalLevelOID" : "3a9aa6d6dbbb10048c65af58217aa75e",
			    "definitionOID" : "3a99fc97dbbb10048c65af58217aa75e",
			    "isMain" : 1,
			    "objectVersion" : 2,
			    "occupantOID" : "3a6b1714dbbb10048c65af58217aa75e",
			    "oid" : "1aee3589dd0110048c83fb23510cb5dd",
			    "organizationUnit" : {
				    "id" : "13120",
				    "manager" : {
					    "cost" : 0,
					    "enableSubstitute" : 1,
					    "id" : "C0160",
					    "identificationType" : "DEFAULT",
					    "ldapid" : "DongTY",
					    "localeString" : "zh_CN",
					    "mailAddress" : "dtianyu@hanbell.cn",
					    "mailingFrequencyType" : 0,
					    "objectVersion" : 15,
					    "oid" : "3a6b1714dbbb10048c65af58217aa75e",
					    "password" : "s//RxFNKLw5LqwI3D34/DWUdH1w=",
					    "phoneNumber" : "65299",
					    "userName" : "董天雨",
					    "workflowServerOID" : "00000000000000WorkflowServer0001"
				    },
				    "objectVersion" : 8,
				    "oid" : "3aab5782dbbb10048c65af58217aa75e",
				    "organizationOID" : "3a8a4bb8dbbb10048c65af58217aa75e",
				    "organizationUnitName" : "维护组",
				    "organizationUnitType" : 0,
				    "superUnitOID" : "3aab8213dbbb10048c65af58217aa75e",
				    "validType" : 1
			    },
			    "specifiedManagerOID" : "3a6e16b1dbbb10048c65af58217aa75e",
			    "users" : {
				    "cost" : 0,
				    "enableSubstitute" : 1,
				    "id" : "C0160",
				    "identificationType" : "DEFAULT",
				    "ldapid" : "DongTY",
				    "localeString" : "zh_CN",
				    "mailAddress" : "dtianyu@hanbell.cn",
				    "mailingFrequencyType" : 0,
				    "objectVersion" : 15,
				    "oid" : "3a6b1714dbbb10048c65af58217aa75e",
				    "password" : "s//RxFNKLw5LqwI3D34/DWUdH1w=",
				    "phoneNumber" : "65299",
				    "userName" : "董天雨",
				    "workflowServerOID" : "00000000000000WorkflowServer0001"
			    }
		    }
	    ]
    }//為符合class格式自己加上去的

    */

    class MANAGER
    {
        public string cost { get; set; }
        public string enableSubstitute { get; set; }
        public string id { get; set; }                  //员工代号id
        public string identificationType { get; set; }
        public string ldapid { get; set; }
        public string localeString { get; set; }
        public string mailAddress { get; set; }
        public string mailingFrequencyType { get; set; }
        public string objectVersion { get; set; }
        public string oid { get; set; }
        public string password { get; set; }
        public string phoneNumber { get; set; }
        public string userName { get; set; }            //员工姓userName
        public string workflowServerOID { get; set; }

        public MANAGER() { }
        
        //取得單身欄位
        public Dictionary<string, string[]> GetDepartmentItem()
        {
            Dictionary<string, string[]> Items = new Dictionary<string, string[]>();

            //需要的
            Items.Add("Columns", new string[] { "id", "userName" });
            Items.Add("Values", new string[] { this.id, this.userName });

            return Items;
        }
    }

    //有用到的在這
    class orgUnit
    {
        public string id { get; set; }                      //部门代号organizationUnit.id
        public string objectVersion { get; set; }
        public string oid { get; set; }
        public string organizationOID { get; set; }
        public string organizationUnitName { get; set; }    //部门名称organizationUnit.organizationUnitName
        public string organizationUnitType { get; set; }
        public string superUnitOID { get; set; }
        public string validType { get; set; }
        public MANAGER manager { get; set; }

        public orgUnit()
        {
            this.manager = new MANAGER();
        }

        //取得單身欄位
        public Dictionary<string, string[]> GetDepartmentItem()
        {
            Dictionary<string, string[]> Items = new Dictionary<string, string[]>();

            //需要的
            Items.Add("Columns", new string[] { "id", "organizationUnitName" });
            Items.Add("Values", new string[] { this.id, this.organizationUnitName });

            return Items;
        }

    }

    class USERS
    {
        public string cost { get; set; }
        public string enableSubstitute { get; set; }
        public string id { get; set; }
        public string identificationType { get; set; }
        public string ldapid { get; set; }
        public string localeString { get; set; }
        public string mailAddress { get; set; }
        public string mailingFrequencyType { get; set; }
        public string objectVersion { get; set; }
        public string oid { get; set; }
        public string password { get; set; }
        public string phoneNumber { get; set; }
        public string userName { get; set; }
        public string workflowServerOID { get; set; }

        public USERS() { }
    }

    class EmployeeDepartmentItem
    {
        public string approvalLevelOID { get; set; }
        public string definitionOID { get; set; }
        public string isMain { get; set; }
        public string objectVersion { get; set; }
        public string occupantOID { get; set; }
        public string oid { get; set; }
        public string specifiedManagerOID { get; set; }
        public orgUnit organizationUnit { get; set; }
        public USERS users { get; set; }

        public EmployeeDepartmentItem()
        {
            this.organizationUnit = new orgUnit();
            this.users = new USERS();
        }
    }

    //部門集合
    class GetEmployeeDepartment
    {
        public List<EmployeeDepartmentItem> EmployeeDepartmentItems { get; set; }

        public GetEmployeeDepartment()
        {
            this.EmployeeDepartmentItems = new List<EmployeeDepartmentItem>();
        }

        //回傳部門table
        public DataTable GetDeparmentTable(string pSearch)
        {
            Dictionary<string, string[]> DepartmentItems = new Dictionary<string, string[]>();  //取得<欄位/值, 陣列值>

            DataTable DepartmentTable = new DataTable();
            for (int i = 0; i < this.EmployeeDepartmentItems.Count; i++)
            {
                bool tInsertData = false;               //是否insert
                DepartmentItems = this.EmployeeDepartmentItems[i].organizationUnit.GetDepartmentItem(); //取得organizationUnit資料

                //第一次進來要新增欄位
                if (i == 0) foreach (var col in DepartmentItems["Columns"]) DepartmentTable.Columns.Add(col);

                //檢查搜尋
                if (string.IsNullOrEmpty(pSearch))
                {
                    tInsertData = true;
                }
                else
                {
                    //有搜尋字段才需處理搜尋
                    foreach (var value in DepartmentItems["Values"])
                    {
                        if (value.Contains(pSearch))
                        {
                            tInsertData = true;//找到就跳出
                            break;
                        }
                    }
                }

                //新增資料
                if (tInsertData) DepartmentTable.Rows.Add(DepartmentItems["Values"]);
            }

            return DepartmentTable;
        }

    }
    #endregion

    #region 加班部門查詢
    /*
     {//為符合class格式自己加上去的
	    "DepartmentItems" : //為符合class格式自己加上去的
         [{
		        "id" : "10000",
		        "manager" : {
			        "cost" : 0,
			        "enableSubstitute" : 1,
			        "id" : "C0002",
			        "identificationType" : "DEFAULT",
			        "ldapid" : "YuZZ",
			        "localeString" : "zh_CN",
			        "mailAddress" : "philipyu@hanbell.cn",
			        "mailingFrequencyType" : 0,
			        "objectVersion" : 15,
			        "oid" : "3a6a95c3dbbb10048c65af58217aa75e",
			        "password" : "wi4Hym6sACoEcRMbn9cMPhcpyBI=",
			        "phoneNumber" : "6118",
			        "userName" : "余昱暄",
			        "workflowServerOID" : "00000000000000WorkflowServer0001"
		        },
		        "objectVersion" : 5,
		        "oid" : "3aab026adbbb10048c65af58217aa75e",
		        "organizationOID" : "3a8a4bb8dbbb10048c65af58217aa75e",
		        "organizationUnitName" : "汉钟",
		        "organizationUnitType" : 0,
		        "validType" : 1
	        }, {
		        "id" : "11000",
		        "manager" : {
			        "cost" : 0,
			        "enableSubstitute" : 1,
			        "id" : "C0002",
			        "identificationType" : "DEFAULT",
			        "ldapid" : "YuZZ",
			        "localeString" : "zh_CN",
			        "mailAddress" : "philipyu@hanbell.cn",
			        "mailingFrequencyType" : 0,
			        "objectVersion" : 15,
			        "oid" : "3a6a95c3dbbb10048c65af58217aa75e",
			        "password" : "wi4Hym6sACoEcRMbn9cMPhcpyBI=",
			        "phoneNumber" : "6118",
			        "userName" : "余昱暄",
			        "workflowServerOID" : "00000000000000WorkflowServer0001"
		        },
		        "objectVersion" : 10,
		        "oid" : "3aabacbedbbb10048c65af58217aa75e",
		        "organizationOID" : "3a8a4bb8dbbb10048c65af58217aa75e",
		        "organizationUnitName" : "总经理室",
		        "organizationUnitType" : 0,
		        "superUnitOID" : "3aab026adbbb10048c65af58217aa75e",
		        "validType" : 1
	        }
        ] 
     } //為符合class格式自己加上去的
     */
    class DepartmentItem
    {
        public string id { get; set; }  //部门编号id
        public string objectVersion { get; set; }
        public string oid { get; set; }
        public string organizationOID { get; set; }
        public string organizationUnitName { get; set; }//部门名称organizationUnitName
        public string organizationUnitType { get; set; }
        public string validType { get; set; }
        public MANAGER manager { get; set; }//與員工部門查詢的MANAGER相同, 這裡不再新增

        public DepartmentItem()
        {
            this.manager = new MANAGER();
        }

        //取得單身欄位
        public Dictionary<string, string[]> GetDepartmentItem()
        {
            Dictionary<string, string[]> Items = new Dictionary<string, string[]>();

            //需要的
            Items.Add("Columns", new string[] { "id", "organizationUnitName" });
            Items.Add("Values", new string[] { this.id, this.organizationUnitName });

            return Items;
        }
    }

    //部門集合
    class GetDepartment
    {
        public List<DepartmentItem> DepartmentItems { get; set; }

        public GetDepartment()
        {
            this.DepartmentItems = new List<DepartmentItem>();
        }

        //回傳部門table
        public DataTable GetDeparmentTable(string pSearch)
        {
            Dictionary<string, string[]> DepaItems = new Dictionary<string, string[]>();  //取得<欄位/值, 陣列值>

            DataTable DepartmentTable = new DataTable();
            for (int i = 0; i < this.DepartmentItems.Count; i++)
            {
                //bool tInsertData = false;               //是否insert//改成一定要下搜尋
                DepaItems = this.DepartmentItems[i].GetDepartmentItem();

                //第一次進來要新增欄位
                if (i == 0) foreach (var col in DepaItems["Columns"]) DepartmentTable.Columns.Add(col);

                //改成一定要下搜尋
                ////檢查搜尋
                //if (string.IsNullOrEmpty(pSearch))
                //{
                //    tInsertData = true;
                //}
                //else
                //{
                //    //有搜尋字段才需處理搜尋
                //    foreach (var value in DepaItems["Values"])
                //    {
                //        if (value.Contains(pSearch))
                //        {
                //            tInsertData = true;//找到就跳出
                //            break;
                //        }
                //    }
                //}

                //新增資料
                //if (tInsertData)//改成一定要下搜尋
                    DepartmentTable.Rows.Add(DepaItems["Values"]);
            }

            return DepartmentTable;
        }

    }
    #endregion

    #region 部門人員查詢


    class GetDepartmentEmployee
    {
        //有用到的:员工代号id、员工姓userName
        public List<MANAGER> DepartmentEmployee { get; set; }   //與員工部門查詢的MANAGER相同, 這裡不再新增

        public GetDepartmentEmployee()
        {
            this.DepartmentEmployee = new List<MANAGER>();
        }

        //回傳部門table
        public DataTable GetDeparmentTable(string pSearch)
        {
            Dictionary<string, string[]> EmployeeItems = new Dictionary<string, string[]>();  //取得<欄位/值, 陣列值>

            DataTable EmployeeTable = new DataTable();
            for (int i = 0; i < this.DepartmentEmployee.Count; i++)
            {
                bool tInsertData = false;               //是否insert
                EmployeeItems = this.DepartmentEmployee[i].GetDepartmentItem();

                //第一次進來要新增欄位
                if (i == 0) foreach (var col in EmployeeItems["Columns"]) EmployeeTable.Columns.Add(col);

                
                //檢查搜尋
                if (string.IsNullOrEmpty(pSearch))
                {
                    tInsertData = true;
                }
                else
                {
                    //有搜尋字段才需處理搜尋
                    foreach (var value in EmployeeItems["Values"])
                    {
                        if (value.Contains(pSearch))
                        {
                            tInsertData = true;//找到就跳出
                            break;
                        }
                    }
                }

                //新增資料
                if (tInsertData) EmployeeTable.Rows.Add(EmployeeItems["Values"]);
            }

            return EmployeeTable;
        }

    }
    #endregion

    #region 立單回傳結果
    class CreateDocResult
    {
        public string code { get; set; }    //回傳代碼
        public string msg { get; set; }     //回傳訊息

        public CreateDocResult() { }
    }
    #endregion
}