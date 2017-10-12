using com.digiwin.Mobile.MCloud.TransferLibrary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Xml.Linq;

namespace HANBELL
{
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    public class Main : System.Web.Services.WebService
    {

        [WebMethod]
        public string Entry(string pM2Pxml)
        {
            #region 設定必要參數
            //P2MXml
            string tP2Mxml = string.Empty;

            //解析M2PXML
            XDocument tM2Pxml = XDocument.Parse(pM2Pxml);
            #endregion

            #region 取得必要參數
            //讀取關鍵字
            string tProductName = DataTransferManager.GetDataValue(tM2Pxml, "Product");//產品名稱
            string tProgramID = DataTransferManager.GetDataValue(tM2Pxml, "ProgramID");//產品代號
            string tServiceName = DataTransferManager.GetDataValue(tM2Pxml, "ServiceName");//service名稱
            #endregion

            new CustomLogger.Logger(tM2Pxml).WriteInfo("M2P : \r\n" + tM2Pxml);

            try
            {
                switch (tProgramID)
                {
                    case "HANBELL01":   //加班申請單
                        switch (tServiceName)
                        {
                            case "BasicSetting":        //頁面初始化
                            case "PageWebService":      //分頁
                            case "SearchWebService":    //搜尋
                                tP2Mxml = new HANBELL01().Get_HANBELL01_BasicSetting(tM2Pxml);
                                break;

                            case "formType_OnBlur":    //加班類別異動
                                tP2Mxml = new HANBELL01().Get_HANBELL01_formType_OnBuler(tM2Pxml);
                                break;

                            case "Company_OP":          //公司別開窗
                                tP2Mxml = new HANBELL01().Get_HANBELL01_Company_OP(tM2Pxml);
                                break;
                            case "company_OnBlur":   //公司別異動
                            case "company_OnClear":
                                tP2Mxml = new HANBELL01().Get_HANBELL01_company_OnBlur(tM2Pxml);
                                break;

                            case "deptno_OP":           //部門開窗
                                tP2Mxml = new HANBELL01().Get_HANBELL01_deptno_OP(tM2Pxml);
                                break;
                            case "deptno_OnBlur":   //部門異動
                            case "deptno_OnClear":
                                tP2Mxml = new HANBELL01().Get_HANBELL01_deptno_OnBlur(tM2Pxml);
                                break;

                            case "Del":     //刪除單身
                                tP2Mxml = new HANBELL01().Get_HANBELL01_Del(tM2Pxml);
                                break;

                            case "CreateDoc": //立單
                                tP2Mxml = new HANBELL01().Get_HANBELL01_CreateDoc(tM2Pxml);
                                break;
                        }
                        break;

                    case "HANBELL01_01": //加班申請單名細
                        switch (tServiceName)
                        {
                            case "BasicSetting":        //頁面初始化
                                tP2Mxml = new HANBELL01().Get_HANBELL01_01_BasicSetting(tM2Pxml);
                                break;

                            case "deptno_OP":        //部門開窗
                                tP2Mxml = new HANBELL01().Get_HANBELL01_01_deptno_OP(tM2Pxml);
                                break;
                            case "deptno_OnBlur":   //部門異動
                            case "deptno_OnClear":
                                tP2Mxml = new HANBELL01().Get_HANBELL01_01_deptno_OnBlur(tM2Pxml);
                                break;

                            case "ids_OP":          //人員開窗
                                tP2Mxml = new HANBELL01().Get_HANBELL01_01_ids_OP(tM2Pxml);
                                break;

                            case "Add":             //新增單身
                            case "Edit":            //修改明細
                                tP2Mxml = new HANBELL01().Get_HANBELL01_01_Add(tM2Pxml);
                                break;
                        }
                        break;
                }
            }
            catch (Exception err)
            {
                return err.Message.ToString();
            }

            new CustomLogger.Logger(tM2Pxml).WriteInfo("P2M : \r\n" + tP2Mxml);

            return tP2Mxml;

        }
    }
}
