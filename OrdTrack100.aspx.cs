using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.IO;
using System.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html;
using iTextSharp.text.html.simpleparser;
using System.Transactions;
using System.Net.Mail;
using System.Data.OleDb;
using System.Diagnostics;


using Excel = Microsoft.Office.Interop.Excel;
using ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat;

public partial class OrdTrack : System.Web.UI.Page
{
    //by yenchang practice-------------------------------------
    //設定全域變數
    public class Global
    {
        //Gridview1 條件
        public static string sqlstring2 = null;
        //
        public static string sqlstring = null;
        public static string sqlstring3 = null;
        public static string sWhere = null;
        public static string sWhere2 = null;
        public static string sRecord = null;
        //追加記錄寄送郵件內容
        public static string sMail_Record = null;
        public static List<String> cb_list = new List<String>();
        //追加記錄rownum與textbox內容,來方便組字串
        public static List<String> lb_rnum_gd2 = new List<String>();
        public static List<String> txComment_gd2 = new List<String>();
        //Delay Order
        public static string sWhere_Delay = null;
        public static string sDelay = null;


    }

 

    protected void Page_Load(object sender, EventArgs e)
    {
        //參考 http://edda.logdown.com/edda/194845-aspnet-confirm-confirmation-window
        Button2.Attributes.Add("onclick ", "return confirm('Do you confirm to send mail?');");
        //Button2.Visible = false;
        Session["web_site"] = Request.CurrentExecutionFilePath.ToString();
        if (!Page.IsPostBack)
        {
            if (Session["user"] == null)
            {

                Session.Clear();
                Response.Write("<script language=javascript>alert('Please Login In!');location.href='Login.aspx';</script>");
            }
            else
            {
                //寫入log寫法
                //Update_Show(Session["user"].ToString());
                Button1_Click(this, e);
                SqlUpdate();
                clear_condition();
                //Session["delay_query"] = "false";
                //if (Session["user"].ToString() == "AA0038" || Session["user"].ToString() == "AAB601" || Session["user"].ToString() == "DA0020")
                for (int i = 0; i < GridView1.PageSize; i++)
                {
                    Session.Remove("page" + i);
                }
                /*
                if (Session["user"].ToString() == "JAP105" || Session["user"].ToString() == "JAP107")
                {
                    //{

                    //Session.Remove("page" + i);
                    Response.Redirect("OrdTrack105.aspx");

                }
                else
                {
                    //Response.Redirect("OrdTrackStandard.aspx");
                }
                 */
                if (Session["user"].ToString().Length >= 6)
                {
                    if (Session["user"].ToString() == "JAP105" || Session["user"].ToString() == "JAP107")
                    {
                        //{

                        //Session.Remove("page" + i);
                        Response.Redirect("OrdTrack105.aspx");

                    }
                    else if (Session["user"].ToString().Substring(0, 6) == "JAP100")
                    {
                        //Response.Redirect("OrdTrack100.aspx");
                    }
                    else
                    {
                        Response.Redirect("OrderTrackStandard.aspx");
                    }
                }
                /*
                //判所client端是否有設定代理伺服器
                string sCus_Ip = "";

                //12142015 By Alex Mark 客戶發生HTTP_X_FORWARDED_FOR為空值的狀況
                //if (Request.ServerVariables["HTTP_VIA"] == null)
                //{
                //    sCus_Ip = Request.ServerVariables["REMOTE_ADDR"].ToString();
                //}
                //else
                //{
                //    sCus_Ip = Request.ServerVariables["HTTP_X_FORWARDED_FOR"].ToString();
                //}

                sCus_Ip = Request.ServerVariables["REMOTE_ADDR"].ToString();
                  
                if (Session["cus_ip"].ToString() == sCus_Ip.ToString())
                {
                    string connectionString = "Data Source=EC2;Initial Catalog=EC2;User ID=EC2;PASSWORD=gmtc";
                    using (SqlConnection conns = new SqlConnection(connectionString))
                    {
                        conns.Open();

                        //設定交易
                        SqlTransaction txn = conns.BeginTransaction();
                        SqlCommand cmds = conns.CreateCommand();
                        cmds.Transaction = txn;

                        try
                        {
                            cmds.CommandText = @"DELETE FROM OTD_ORD_TMP
                                                  WHERE ACCOUNT = '" + Session["user"] + @"' 
                                                        AND CUS_IP = '" + Session["cus_ip"] + "'";

                            cmds.ExecuteNonQuery();
                            cmds.Parameters.Clear();

                            cmds.CommandText = @"INSERT INTO OTD_RECORD (CUS_ID, IP_ADDRESS, IP_TIME, KIND_TYPE)
                                                      VALUES ('" + Session["user"] + "', '" + sCus_Ip + @"', GETDATE(), 'OTD SERVICE CLICK')";
     
                            cmds.ExecuteNonQuery();
                            cmds.Parameters.Clear();

                            txn.Commit();

                        }
                        catch (Exception ex)
                        {
                            txn.Rollback();
                            //Response.Write("Otdlist.aspx");
                            //Response.Write(ex.Message);
                        }
                    }    
                }
            //}
            /*else
            {
                Response.Redirect("IndexEn.aspx");
            }
            */
            }
        }
        else
        {
            //Session["delay_query"] = "false";
        }
        //Global.sDelay = "";
    }

    //寄送郵件
    
    /*protected void Button2_Click(object sender, EventArgs e)
    {
        string sMail_content = "";
        /*
        string connectionString = "Data Source=EC2;Initial Catalog=EC2;User ID=EC2;PASSWORD=gmtc";

        using (SqlConnection conns = new SqlConnection(connectionString))
        {
            conns.Open();

            //設定交易
            SqlTransaction txn = conns.BeginTransaction();
            SqlCommand cmds = conns.CreateCommand();
            cmds.Transaction = txn;

            try
            {
                cmds.CommandText = @"INSERT INTO OTD_ORD (GRADE, SIZE_I, QTY_LB, ETA, EX_MILL_DATE, ACCOUNT, ORD_DATE, CUS_IP)
                                          SELECT GRADE, SIZE_I, QTY_LB, ETA, EX_MILL_DATE, ACCOUNT, GETDATE(), CUS_IP
                                            FROM OTD_ORD_TMP
                                           WHERE ACCOUNT = '" + Session["user"] + @"'
                                                 AND CUS_IP = '" + Session["cus_ip"] + "'";

                cmds.ExecuteNonQuery();
                cmds.Parameters.Clear();


                string mail_content = "";
                cmds.CommandText = @"SELECT *
                                       FROM OTD_ORD_TMP
                                      WHERE ACCOUNT = '" + Session["user"] + @"'
                                            AND CUS_IP = '" +Session["cus_ip"] + "'";

                using (SqlDataReader odr = cmds.ExecuteReader())
                {
                    while (odr.Read())
                    {
                        mail_content = mail_content + "\n" + "GRADE:" + odr["GRADE"] + ", SIZE(inch):" + odr["SIZE_I"] + ", Type:" + odr["TYPE"] + ", Qty:" + odr["QTY_LB"] + ", EX MILL DATE:" + Convert.ToDateTime(odr["EX_MILL_DATE"]).ToShortDateString() + ", ETA:" + Convert.ToDateTime(odr["ETA"]).ToShortDateString();
                    }
                }

                mail_content = mail_content + "\n\n" + "Comment: \n" + TextBox4.Text;


                string sBrief_Nm = "";
                string sCustomer_Email = "";
                string sSalesmn = "";
                string sSale_Email = "";
                string sSale_Mgr_Email = "";

                using (SqlConnection conns1 = new SqlConnection(connectionString))
                {
                    conns1.Open();
                    SqlCommand cmds1 = conns1.CreateCommand();

                    cmds1.CommandText = @"SELECT A.BRIEF_NM, A.EMAIL, A.SALESMN, B.SALE_EMAIL, B.SALE_MGR_EMAIL
                                            FROM PAL_EC_APPLY AS A INNER JOIN
                                                 PAL_SALE AS B ON A.SALESMN = B.SALE
                                           WHERE (A.UNI_NO = '" + Session["user"] + "')";

                    using (SqlDataReader odr1 = cmds1.ExecuteReader())
                    {
                        while (odr1.Read())
                        {
                            sBrief_Nm = odr1["BRIEF_NM"].ToString();
                            sSalesmn = odr1["SALESMN"].ToString();
                            sSale_Email = odr1["SALE_EMAIL"].ToString();
                            sSale_Mgr_Email = odr1["SALE_MGR_EMAIL"].ToString();
                            sCustomer_Email = odr1["EMAIL"].ToString();
                        }
                    }
                }

                //先將收件人組合起來
                string sReceiver = "";
                //sReceiver = sSale_Email + "," + sSale_Mgr_Email + ",alex.yang@gmtc.com.tw, yltsai@gmtc.com.tw," + sCustomer_Email;
                //04112016 By Alex Edit 延林課長為了DEMO，所以先取消寄信給業務
                sReceiver = sSale_Email + "," + sSale_Mgr_Email + ",alex.yang@gmtc.com.tw, yltsai@gmtc.com.tw,";

                //設定SMTP
                System.Net.Mail.SmtpClient MySmtp = new System.Net.Mail.SmtpClient("192.168.1.1", 25);

                //發送Email("寄信者","收信者","主題","內容")
                //MySmtp.Send("admin@gmtc.com.tw", "alex.yang@gmtc.com.tw", "OTD order test", @"Dear 您好:" + "\r\n" + Session["user"] + " Order OTD List. Please remind customer send P/O for you. Thank you." + "\r\n " + mail_content);
                MySmtp.Send("admin@gmtc.com.tw", sReceiver, "OTD order test", "Dear " + sSalesmn + " :" + "\r\n" + sBrief_Nm + " Order OTD product. Please remind customer send P/O for you. Thank you." + "\r\n " + mail_content);
                Response.Write("<script language=javascript>alert('Send Mail Successed!!');location.href='OtdList.aspx';</script>");

                cmds.CommandText = @"DELETE FROM OTD_ORD_TMP
                                           WHERE ACCOUNT = '" + Session["user"] + @"'
                                                 AND CUS_IP = '" + Session["cus_ip"] + "'";

                cmds.ExecuteNonQuery();
                cmds.Parameters.Clear();

                txn.Commit();

                GridView2.DataBind();
            }
            catch (Exception ex)
            {
                txn.Rollback();
                Response.Write("Otdlist.aspx");
                //Response.Write(ex.Message);
            }
        }   
        
        string connectionString = "Data Source=GHICNEOR.WORLD;User ID=sd;PASSWORD=sd";
        using (OracleConnection conns = new OracleConnection(connectionString))
        {
            conns.Open();


            //設定查詢
            OracleCommand cmds = conns.CreateCommand();
            //cmds.CommandText = "SELECT COUNT(*) AS DCOUNT FROM OV_CERT_DOWNLOAD WHERE CUS_ID = @CUS_ID " + Global.sqlstring2;
            cmds.CommandText = @"select a.po_no,c.order_no,c.order_seq,c.wip_lot from cop.ord_mst a
                      ,cop.ord_dt b,wip.ship_plan c 
                      where a.cus_id= '" + Session["user"] + "' " + Global.sqlstring2 + Global.sWhere +
                @" and a.ord_id=b.ord_id and a.ctrl='Y' and b.ord_status<>'C' and c.order_type='NEW' 
                      and b.ord_id=to_number(c.order_no) and b.ord_seq=to_number(c.order_seq)
                      union
                      select a.po_no,c.order_no,c.order_seq,c.wip_lot
                      from cop.ord_mst a,cop.ord_dt b,wip.ship_plan_load c 
                      where a.cus_id= '" + Session["user"] + "' " + Global.sqlstring2 + Global.sWhere +
                @" and a.ord_id=b.ord_id and a.ctrl='Y' and 
                      b.ord_status<>'C' and c.order_type='NEW' 
                      and b.ord_id=to_number(c.order_no) and b.ord_seq=to_number(c.order_seq) ";
            //OracleParameter param = new OracleParameter(":CUS_ID", OracleDbType.Varchar2);
            //param.Value = Session["user"];
            //cmds.Parameters.Add(param);

            using (OracleDataReader ora = cmds.ExecuteReader())
            {

                while (ora.Read())
                {
                    //Label6.Text = odr["total"].ToString();
                    sMail_content = sMail_content + "\n" + ora["po_no"] + "\t" + ora["order_no"] + "\t" + ora["order_seq"] + "\t" + ora["wip_lot"];
                }
            }
        }
        //先將收件人組合起來
        string sReceiver = "";
        sReceiver = "YenChang.Su@gmtc.com.tw";
        //設定SMTP
        System.Net.Mail.SmtpClient MySmtp = new System.Net.Mail.SmtpClient("192.168.1.1", 25);
        sMail_content = TextBox4.Text + "\n\n\n" + sMail_content;
        //發送Email("寄信者","收信者","主題","內容")
        //MySmtp.Send("admin@gmtc.com.tw", "alex.yang@gmtc.com.tw", "OTD order test", @"Dear 您好:" + "\r\n" + Session["user"] + " Order OTD List. Please remind customer send P/O for you. Thank you." + "\r\n " + mail_content);
        MySmtp.Send("admin@gmtc.com.tw", sReceiver, "Ord Track Test", "Dear \r\n  Order Track list. Please remind customer send P/O for you. Thank you." + "\r\n\n\n " + sMail_content);
        Response.Write("<script language=javascript>alert('Send Mail Successed!!');location.href='Ordtrack.aspx';</script>");


    }
    */
    private void PeopleCount()
    {
        //參考網址https://dotblogs.com.tw/wesley0917/2010/12/26/20387
        //如果個人變數是第一次產生時。
        //if (Session.IsNewSession)
        if (Session["user"].ToString() == "JAP100")
        {
            //如果全域變數等於空值時。
            if (Application["PageCounter_A"] == null)
            {
                //鎖定全域變數存取。
                Application.Lock();
                //指派全域變數初始值。
                Application["PageCounter_A"] = 0;
                //不鎖定全域變數存取。
                Application.UnLock();
            }
            //指派全域變數累加值。
            Application.Set("PageCounter_A", (Convert.ToInt32(Application["PageCounter_A"]) + 1));
            //將全域變數值指派給個人變數值。
            Session["Visited_S"] = Application["PageCounter_A"];
        }
        //輸出結果。
        //Response.Write("你是第 " + Session["Visited_S"] + " 位訪客！");
        Labe20.Text=Session["Visited_S"].ToString();
        
    }
    //記憶分頁checkbox控制項的fun
    /*
    protected void GridView1_PreRender(object sender, EventArgs e)
    {
        //Global.sWhere = null;
        //Session["rownum"] = rownum_box;
        //Session["checked"] = checked_box;
        if (Session["rownum"] != null)
        {
            CheckBox chb;　
            CheckBox ChkBoxHeader = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");
            bool[] values = (bool[])Session["page" + GridView1.PageIndex];
            List<bool> checked_box = (List<bool>)Session["checked"];
            List<string> rownum_box = (List<string>)Session["rownum"];
            for (int i = 0; i < GridView1.Rows.Count; i++)
            {
                chb = (CheckBox)GridView1.Rows[i].FindControl("CheckBox1");
                Label rownum = (Label)GridView1.Rows[i].FindControl("Label7");
                int index = rownum_box.IndexOf(rownum.Text);
                //if (values[i] == true)
                //bool chb_checked = checked_box.IndexOf(rownum_box[i]);
                //bool chb_checked = checked_box[index];
                if (index != -1)
                {
                    chb.Checked = checked_box[index];
                }//}
                //else if (HiddenField4.Value == "true")
                /*
                if (HiddenField4.Value == "true")
                {
                    ChkBoxHeader.Checked = true;
                    chb.Checked = true;
                    //CheckboxallRecord();
                    //ContentQry();
                }
                */
                /*
                if (HiddenField4.Value == "true")
                {
                    ChkBoxHeader.Checked = true;
                    chb.Checked = true;
                    CheckboxallRecord();
                    ContentQry();
                }
                else
                {
                    ChkBoxHeader.Checked = false;
                    chb.Checked = false;
                    CheckboxallRecord();
                    ContentQry();
                }
                */
            //}
            /*
            if (HiddenField4.Value == "true")
            {
                ChkBoxHeader.Checked = true;
            }
            else
            {
                ChkBoxHeader.Checked = false;
            }
            */
       // }

    //}
    protected void GridView1_PreRender(object sender, EventArgs e)
    {
        //Global.sWhere = null;
        if (Session["page" + GridView1.PageIndex] != null)
        {
            CheckBox chb;
            bool[] values = (bool[])Session["page" + GridView1.PageIndex];
            for (int i = 0; i < GridView1.Rows.Count; i++)
            {
                chb = (CheckBox)GridView1.Rows[i].FindControl("CheckBox1");
                chb.Checked = values[i];

            }
        }
        //控制項需寫到最後面
        if (CheckBox4.Checked ==true)
        {
            CheckedAllBox();
        }
        else if (CheckBox4.Checked == false)
        {
            CheckedAllBox();
        }
        /*
        if (HiddenField4.Value == "true")
        {
            CheckBox ChkBoxHeader = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");
            ChkBoxHeader.Checked = true;
        }
        */

    }

    //記錄GridView1 每個page勾選的checkbox並把該列rownum植入Global.cb_list
    /*
    private void CheckBoxRecord()
    {
        #region 記錄被勾選的Checkbox
        //Reference:http://forums.asp.net/t/1147075.aspx?Gridview+CheckBox+Persist+in+Paging(11/12/2015)
        int d = GridView1.PageCount;
        //bool[] values = new bool[GridView1.PageSize];
        //List<bool> checked_box = new List<bool>();
        //List<string> rownum_box = new List<string>();
        //CheckBox chb;
        //Label rownum;
        bool[] values = new bool[GridView1.PageSize];
        Global.sWhere = null;
        for (int i = 0; i < GridView1.Rows.Count; i++)
        {
            //chb = (CheckBox)GridView1.Rows[i].FindControl("CheckBox1");
            //rownum = (Label)GridView1.Rows[i].FindControl("Label7");

            //if (chb != null&&rownum!=null)
            //{
                checked_box.Add(chb.Checked);
                rownum_box.Add(rownum.Text);
                if (checked_box[i] == true)
                {
                    #region 記錄哪列改變
                    //CheckBox cb = (CheckBox)sender;
                    GridViewRow row = (GridViewRow)chb.NamingContainer;
                    #endregion
                    //Label lb6 = (row.Cells[2].FindControl("Label2") as Label);
                    Label lb3 = (row.Cells[24].FindControl("Label7") as Label);
                    if (Global.cb_list.Count == 0)
                    {
                        Global.cb_list.Add(lb3.Text.ToString());
                    }
                    else
                    {
                        bool add = true;
                        for (int k = 0; k < Global.cb_list.Count; k++)
                        {
                            if (Global.cb_list[k] == lb3.Text.ToString())
                                add = false;
                        }
                        if (add)
                        {
                            Global.cb_list.Add(lb3.Text.ToString());
                        }
                    }
                    //Global.cb_list.Add(lb3.Text.ToString());

                }
                else
                {
                    #region 記錄哪列改變
                    //CheckBox cb = (CheckBox)sender;
                    GridViewRow row = (GridViewRow)chb.NamingContainer;
                    Label lb3 = (row.Cells[24].FindControl("Label7") as Label);
                    #endregion
                    /*
                    bool del = false;
                    int del_index = 0;
                    for (int k = 0; k < Global.cb_list.Count; k++)
                    {
                        if (Global.cb_list[k] == lb3.Text.ToString())
                        {
                            del = true;
                            del_index = k;
                        }
                    }
                    if (del)
                    {

                        Global.cb_list.RemoveAt(del_index);
                    }
                    /*
                    if (lb3.Text.ToString() != null)
                    {
                        if (Global.cb_list.Contains(lb3.Text.ToString()))
                        {
                            Global.cb_list.Remove(lb3.Text.ToString());
                        }
                    }
                    

                }


            }

        }
            //將記錄的rownum與checked情形記錄
            Session["rownum"] = rownum_box;
            Session["checked"] = checked_box;
        
        #endregion
    }
    */
    private void CheckBoxRecord()
    {
        #region 記錄被勾選的Checkbox
        //Reference:http://forums.asp.net/t/1147075.aspx?Gridview+CheckBox+Persist+in+Paging(11/12/2015)
        int d = GridView1.PageCount;
        bool[] values = new bool[GridView1.PageSize];
        CheckBox chb;
        Global.sWhere = null;
        for (int i = 0; i < GridView1.Rows.Count; i++)
        {
            chb = (CheckBox)GridView1.Rows[i].FindControl("CheckBox1");

            if (chb != null)
            {
                values[i] = chb.Checked;
                if (values[i] == true)
                {
                    #region 記錄哪列改變
                    //CheckBox cb = (CheckBox)sender;
                    GridViewRow row = (GridViewRow)chb.NamingContainer;
                    #endregion
                    //Label lb6 = (row.Cells[2].FindControl("Label2") as Label);
                    Label lb3 = (row.Cells[24].FindControl("Label7") as Label);
                    if (Global.cb_list.Count == 0)
                    {
                        Global.cb_list.Add(lb3.Text.ToString());
                    }
                    else
                    {
                        bool add = true;
                        for (int k = 0; k < Global.cb_list.Count; k++)
                        {
                            if (Global.cb_list[k] == lb3.Text.ToString())
                                add = false;
                        }
                        if (add)
                        {
                            Global.cb_list.Add(lb3.Text.ToString());
                        }
                    }
                    //Global.cb_list.Add(lb3.Text.ToString());

                }
                else
                {
                    #region 記錄哪列改變
                    //CheckBox cb = (CheckBox)sender;
                    GridViewRow row = (GridViewRow)chb.NamingContainer;
                    Label lb3 = (row.Cells[24].FindControl("Label7") as Label);
                    #endregion
                    /*
                    bool del = false;
                    int del_index = 0;
                    for (int k = 0; k < Global.cb_list.Count; k++)
                    {
                        if (Global.cb_list[k] == lb3.Text.ToString())
                        {
                            del = true;
                            del_index = k;
                        }
                    }
                    if (del)
                    {

                        Global.cb_list.RemoveAt(del_index);
                    }
                    */
                    if (lb3.Text.ToString() != null)
                    {
                        if (Global.cb_list.Contains(lb3.Text.ToString()))
                        {
                            Global.cb_list.Remove(lb3.Text.ToString());
                        }
                    }
                    

                }


            }

        }

        Session["page" + GridView1.PageIndex] = values;

        #endregion
    }


    //查詢
    protected void Button1_Click(object sender, EventArgs e)
    {
        try
        {
            //Global.cb_list.Clear();
            //Global.lb_rnum_gd2.Clear();
            //Global.txComment_gd2.Clear();
            Global.sDelay= "false";
            Global.sqlstring2 = null;
            Global.sRecord = null;
            //暫時取消全選功能 by yenchang 07/13/2016 
            //CheckBox ChkBoxHeader = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");
            //ChkBoxHeader.Checked = false;
            HiddenField4.Value = "false"; 
            //if((DropDownList)GridView1.HeaderRow.FindControl("DropDownList1")).Cells[1].ToString().Trim()=="COMMON")
            if (((DropDownList)GridView1.HeaderRow.FindControl("DropDownList1")).SelectedValue == "COMMON")
            {
                //Global.sqlstring2 += " AND A.PO_NO BETWEEN '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox4")).Text.ToString().Trim() + "' AND '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox5")).Text.ToString().Trim() + "'";
                //HiddenField1.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text.ToString().Trim();
                //HiddenField9.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox18")).Text.ToString().Trim();
                Global.sqlstring2 += " AND TOSHIBA_PROJECT LIKE '%" + ((DropDownList)GridView1.HeaderRow.FindControl("DropDownList1")).SelectedValue + "%'";
                HiddenField8.Value = ((DropDownList)GridView1.HeaderRow.FindControl("DropDownList1")).SelectedValue;
                Global.sRecord = Global.sRecord + "TOSHIBA_PROJECT=" + ((DropDownList)GridView1.HeaderRow.FindControl("DropDownList1")).SelectedValue+";";
            }
            else if (((DropDownList)GridView1.HeaderRow.FindControl("DropDownList1")).SelectedValue == "PROJECT")
            {
                Global.sqlstring2 += " AND TOSHIBA_PROJECT NOT LIKE '%COMMON%'";
                //HiddenField1.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim();
                HiddenField8.Value = ((DropDownList)GridView1.HeaderRow.FindControl("DropDownList1")).SelectedValue;
                Global.sRecord = Global.sRecord + "TOSHIBA_PROJECT=" + ((DropDownList)GridView1.HeaderRow.FindControl("DropDownList1")).SelectedValue + ";";
            }
            else 
            {
                Global.sqlstring2 += " ";
                HiddenField8.Value = ((DropDownList)GridView1.HeaderRow.FindControl("DropDownList1")).SelectedValue;
                //Global.sRecord = Global.sRecord + "TOSHIBA_PROJECT=" + ((DropDownList)GridView1.HeaderRow.FindControl("DropDownList1")).SelectedValue + ";";
            } 

            if (((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim() != "")
            {
                //Global.sqlstring2 += " AND A.PO_NO BETWEEN '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox4")).Text.ToString().Trim() + "' AND '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox5")).Text.ToString().Trim() + "'";
                //HiddenField1.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text.ToString().Trim();
                //HiddenField9.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox18")).Text.ToString().Trim();
                Global.sqlstring2 += " AND TOSHIBA_PROJECT LIKE '%" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim() + "%'";
                HiddenField1.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim();
                //Global.sqlstring2 += " AND PO_NO = '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox4")).Text.ToString().Trim()+"'";
                Global.sRecord = Global.sRecord + "TOSHIBA_PROJECT=" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim() + ";";
            }
            else
            {
                //Global.sqlstring2 += " AND A.PO_NO LIKE '%" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox4")).Text.ToString().Trim() + "%'";
                HiddenField1.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim();
                //HiddenField1.Value = ((DropDownList)GridView1.HeaderRow.FindControl("DropDownList1")).SelectedValue;

            }
            if (((TextBox)GridView1.HeaderRow.FindControl("TextBox9")).Text.ToString().Trim() != "")
            {
                //Global.sqlstring2 += " AND A.PO_NO BETWEEN '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox6")).Text.ToString().Trim() + "' AND '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox7")).Text.ToString().Trim() + "'";
                //HiddenField1.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text.ToString().Trim();
                //HiddenField9.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox18")).Text.ToString().Trim();
                Global.sqlstring2 += " AND FUJI_PO LIKE '%" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox9")).Text.ToString().Trim() + "%'";
                HiddenField2.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox9")).Text.ToString().Trim();
                Global.sRecord = Global.sRecord + "FUJI_PO=" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox9")).Text.ToString().Trim() + ";";

            }
            else
            {
                //Global.sqlstring2 += " AND A.PO_NO LIKE '%" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox6")).Text.ToString().Trim() + "%'";
                HiddenField2.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox9")).Text.ToString().Trim();

            }
            if (((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text.ToString().Trim() != "")
            {
                //Global.sqlstring2 += " AND A.PO_NO BETWEEN '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim() + "' AND '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox9")).Text.ToString().Trim() + "'";
                //HiddenField1.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text.ToString().Trim();
                //HiddenField9.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox18")).Text.ToString().Trim();
                Global.sqlstring2 += " AND GRADE LIKE '%" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text.ToString().Trim() + "%'";
                HiddenField3.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text.ToString().Trim();
                Global.sRecord = Global.sRecord + "GRADE=" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text.ToString().Trim() + ";";
            }
            else
            {
                //Global.sqlstring2 += " AND A.PO_NO LIKE '%" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim() + "%'";
                HiddenField3.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text.ToString().Trim();

            }
            if (((TextBox)GridView1.HeaderRow.FindControl("TextBox11")).Text.ToString().Trim() != "")
            {
                //Global.sqlstring2 += " AND A.PO_NO BETWEEN '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim() + "' AND '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox9")).Text.ToString().Trim() + "'";
                //HiddenField1.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text.ToString().Trim();
                //HiddenField9.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox18")).Text.ToString().Trim();
                Global.sqlstring2 += " AND TSB_PO LIKE '%" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox11")).Text.ToString().Trim() + "%'";
                HiddenField5.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox11")).Text.ToString().Trim();
                Global.sRecord = Global.sRecord + "TSB_PO=" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox11")).Text.ToString().Trim() + ";";

            }
            else
            {
                //Global.sqlstring2 += " AND A.PO_NO LIKE '%" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim() + "%'";
                HiddenField5.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox11")).Text.ToString().Trim();

            }
            if (((TextBox)GridView1.HeaderRow.FindControl("TextBox12")).Text.ToString().Trim() != "" && ((TextBox)GridView1.HeaderRow.FindControl("TextBox13")).Text.ToString().Trim() != "")
            {
                //Global.sqlstring2 += " AND A.PO_NO BETWEEN '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim() + "' AND '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox9")).Text.ToString().Trim() + "'";
                //HiddenField1.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text.ToString().Trim();
                //HiddenField9.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox18")).Text.ToString().Trim();
                //Global.sqlstring2 += " AND EX_WORKS >= " + Convert.ToDateTime(((TextBox)GridView1.HeaderRow.FindControl("TextBox12")).Text.ToString().Trim()) + "AND EX_WORKS <= " + Convert.ToDateTime(((TextBox)GridView1.HeaderRow.FindControl("TextBox13")).Text.ToString().Trim());
                //AjaxControlToolkit.CalendarExtender calenderDate = new AjaxControlToolkit.CalendarExtender();
                //DateTime temp = Convert.ToDateTime(((TextBox)GridView1.HeaderRow.FindControl("TextBox12")).Text.ToString());
                Global.sqlstring2 += " AND EX_WORKS BETWEEN '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox12")).Text.ToString().Trim() + "' AND '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox13")).Text.ToString().Trim() + "'";
                HiddenField6.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox12")).Text.ToString().Trim();
                HiddenField7.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox13")).Text.ToString().Trim();
                Global.sRecord = Global.sRecord + "EX_WORKS BETWEEN " + ((TextBox)GridView1.HeaderRow.FindControl("TextBox12")).Text.ToString().Trim() + " AND " + ((TextBox)GridView1.HeaderRow.FindControl("TextBox13")).Text.ToString().Trim() + ";";

            }
            else
            {
                //Global.sqlstring2 += " AND A.PO_NO LIKE '%" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim() + "%'";
                Response.Write("<script language=javascript>alert('Please Input EX WORKS! Format = 'MM/DD/YYYY' ');</script>");
                HiddenField6.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox12")).Text.ToString().Trim();
                HiddenField7.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox13")).Text.ToString().Trim();

            }
            if (((TextBox)GridView1.HeaderRow.FindControl("TextBox67")).Text.ToString().Trim() != "" && ((TextBox)GridView1.HeaderRow.FindControl("TextBox68")).Text.ToString().Trim() != "")
            {
                //Global.sqlstring2 += " AND A.PO_NO BETWEEN '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim() + "' AND '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox9")).Text.ToString().Trim() + "'";
                //HiddenField1.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text.ToString().Trim();
                //HiddenField9.Value = ((TextBox)GridView1.HeaderRow.FindControl("TextBox18")).Text.ToString().Trim();
                //Global.sqlstring2 += " AND EX_WORKS >= " + Convert.ToDateTime(((TextBox)GridView1.HeaderRow.FindControl("TextBox12")).Text.ToString().Trim()) + "AND EX_WORKS <= " + Convert.ToDateTime(((TextBox)GridView1.HeaderRow.FindControl("TextBox13")).Text.ToString().Trim());
                //AjaxControlToolkit.CalendarExtender calenderDate = new AjaxControlToolkit.CalendarExtender();
                //DateTime temp = Convert.ToDateTime(((TextBox)GridView1.HeaderRow.FindControl("TextBox12")).Text.ToString());
                Global.sqlstring2 += " AND TSB_NEED_DATE BETWEEN '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox67")).Text.ToString().Trim() + "' AND '" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox68")).Text.ToString().Trim() + "'";
                Session["TSB_NEED_DATE_1"] = ((TextBox)GridView1.HeaderRow.FindControl("TextBox67")).Text.ToString().Trim();
                Session["TSB_NEED_DATE_2"] = ((TextBox)GridView1.HeaderRow.FindControl("TextBox68")).Text.ToString().Trim();
                Global.sRecord = Global.sRecord + "TSB_NEED_DATE BETWEEN " + ((TextBox)GridView1.HeaderRow.FindControl("TextBox67")).Text.ToString().Trim() + " AND " + ((TextBox)GridView1.HeaderRow.FindControl("TextBox68")).Text.ToString().Trim() + ";";

            }
            else
            {
                //Global.sqlstring2 += " AND A.PO_NO LIKE '%" + ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text.ToString().Trim() + "%'";
                Response.Write("<script language=javascript>alert('Please Input TSB_NEED_DATE! Format = 'MM/DD/YYYY' ');</script>");
                Session["TSB_NEED_DATE_1"] = ((TextBox)GridView1.HeaderRow.FindControl("TextBox67")).Text.ToString().Trim();
                Session["TSB_NEED_DATE_2"] = ((TextBox)GridView1.HeaderRow.FindControl("TextBox68")).Text.ToString().Trim();

            }

            String ord_track = @"select TOSHIBA_PROJECT,GMTC_PO,ITEM,FUJI_PO,GRADE,SIZE_,TSB_PO,EX_WORKS,ORD_QTY,SHIPEED_QTY,IN_STOCK_QTY,WIP_QTY,IN_STOCK_PCS,PLAN_TO_SHIP_FROM_GMTC,CSD,
            ETD,ETA,TSB_NEED_DATE,SHIPPING_STATUS,WIP,WEIGHT_TOLERANCE,DELAY,NOTE,ROW_NUM,CUS_ID,WIP_LOT from dbo.OV_ORDER_TRACE_JAP100 where CUS_ID = substring('" + Session["user"] + @"',1,6)" + Global.sqlstring2 + " AND ROW_NUM <> '0' ORDER BY EX_WORKS ";
            
            
            SqlDataSource1.SelectCommand = ord_track;
            GridView1.DataBind();
            //3.ToExcel測試用
            SqlDataSource3.SelectCommand = ord_track;
            SqlDataSource1.SelectParameters.Clear();
            SqlDataSource3.SelectParameters.Clear();
            Query_Record();
            //GridView1.DataBind();
            CheckedBox();
            CheckBoxRecord();
            CheckboxallRecord();
            ContentQry();
            //查詢完先清掉所有CHECKBOX的SESSION BY YENCHANG 07/13/2016
            for (int i = 0; i < GridView1.PageCount; i++)
            {
                Session["page" + i] = null;//清除pageSESSION
            }
            if (CheckBox4.Checked == true)
            {
                Button4_Click(this, e);
            }
            else if (CheckBox5.Checked == true)
            {
                Button5_Click(this, e);
            }
            //GridView3.DataBind();
        }
        catch
        {
            Response.Write("<script language=javascript>alert('Please Input Filter!');location.href='OrdTrack100.aspx';</script>");

        }
    }
    //GridView切換分頁fun
    protected void GridView1_PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        //保留原搜尋條件
        ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text = HiddenField1.Value;
        ((TextBox)GridView1.HeaderRow.FindControl("TextBox9")).Text = HiddenField2.Value;
        ((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text = HiddenField3.Value;
        ((TextBox)GridView1.HeaderRow.FindControl("TextBox11")).Text = HiddenField5.Value;
        ((TextBox)GridView1.HeaderRow.FindControl("TextBox12")).Text = HiddenField6.Value;
        ((TextBox)GridView1.HeaderRow.FindControl("TextBox13")).Text = HiddenField7.Value;
        ((DropDownList)GridView1.HeaderRow.FindControl("DropDownList1")).SelectedValue = HiddenField8.Value;
        if (Session["TSB_NEED_DATE_1"] != null || Session["TSB_NEED_DATE_2"] != null)
        {
            ((TextBox)GridView1.HeaderRow.FindControl("TextBox67")).Text = Session["TSB_NEED_DATE_1"].ToString();
            ((TextBox)GridView1.HeaderRow.FindControl("TextBox68")).Text = Session["TSB_NEED_DATE_2"].ToString();
        }
        GridView1.PageIndex = e.NewPageIndex;
        #region 記錄被勾選的Checkbox
        //Reference:http://forums.asp.net/t/1147075.aspx?Gridview+CheckBox+Persist+in+Paging(11/12/2015)
        int d = GridView1.PageCount;
        bool[] values = new bool[GridView1.PageSize];
        #endregion
        String ord_track = @"select TOSHIBA_PROJECT,GMTC_PO,ITEM,FUJI_PO,GRADE,SIZE_,TSB_PO,EX_WORKS,ORD_QTY,SHIPEED_QTY,IN_STOCK_QTY,WIP_QTY,IN_STOCK_PCS,PLAN_TO_SHIP_FROM_GMTC,CSD,
        ETD,ETA,TSB_NEED_DATE,SHIPPING_STATUS,WIP,WEIGHT_TOLERANCE,DELAY,NOTE,ROW_NUM,CUS_ID,WIP_LOT from dbo.OV_ORDER_TRACE_JAP100 where cus_id = substring('" + Session["user"] + @"',1,6)" + Global.sqlstring2+" AND ROW_NUM <> '0' ORDER BY EX_WORKS";
        SqlDataSource1.SelectCommand = ord_track;
        //CheckBox ChkBoxHeader = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");
        //GridView1.DataBind();



       // Session["page" + GridView1.PageIndex] = values;
        //切分頁記錄
        SqlDataSource1.SelectCommand = ord_track;
        //全選CHECKBOX記錄
        //CheckedAllBox();
        GridView1.DataBind();
        //全選

        
    }



    // 選取gridview 1 checkbox的fun by YenChang
    protected void CheckBox1_CheckedChanged1(object sender, EventArgs e)
    {
        Global.sWhere = null;
        CheckBoxRecord();
        
        if (Global.cb_list.Count != 0)
        {
            Button2.Visible = true;
            Label6.Visible = true;
            DropDownList2.Visible = true;
            GridView2.Visible = true;
            //HiddenField4.Value = "true";
            for (int i = 0; i < Global.cb_list.Count; i++)
            {
                if (Global.sWhere == null)
                {
                    Global.sWhere = " AND ( ROW_NUM = " + Global.cb_list[i] + ")";
                }

                else
                {
                    Global.sWhere = Global.sWhere.Substring(0, Global.sWhere.Length - 1) + " or ROW_NUM = " + Global.cb_list[i] + ")";
                }
            }
        }
        else
        {
            Global.sWhere = " AND ( ROW_NUM =-1)";
            Button2.Visible = false;
        }
        
        //Global.cb_list
        //try
        //{


        //String sqlstr = @"select aa.PO_NO,aa.ORDER_NO,aa.ORDER_SEQ,aa.CERT_PO,aa.WIP_LOT from (select PO_NO,ORDER_NO,ORDER_SEQ,CERT_PO,WIP_LOT,row_number() over (order by PO_NO asc) as [Row_Number] from dbo.OV_ORDER_TRACE_MAIN where cus_id ='JAP100') aa " + Global.sqlstring2 + Global.sWhere;

        String sqlstr = @"select TOSHIBA_PROJECT,GMTC_PO,ITEM,WIP_LOT,ROW_NUM,FUJI_PO,TSB_PO from dbo.OV_ORDER_TRACE_JAP100 where cus_id ='" + Session["user"].ToString().Substring(0, 6) + @"'" + Global.sqlstring2 + Global.sWhere + " AND ROW_NUM <> '0'";   
        SqlDataSource2.SelectCommand = sqlstr;
        //GridView2.DataBind();
        //Global.lb_rnum_gd2.Clear();
        //Global.txComment_gd2.Clear();
        CommentRecord();
        GridView2.DataBind();
        GridView2.Visible = true;
        Button2.Focus();
        //catch (Exception ex)
        //{
        //    Response.Redirect("Ordtrack.aspx");
        //}
    }
    protected void LinkButton2_Click(object sender, EventArgs e)
    {
        #region 記錄哪列改變
        LinkButton pb = (LinkButton)sender;
        GridViewRow row = (GridViewRow)pb.NamingContainer;
        #endregion
        Label pblabel = (row.Cells[6].FindControl("Label5") as Label);
        Session["gmtc_po"] = pblabel.Text;
        this.Response.Write("<script language=javascript>window.open('Detail_Track.aspx','Detail','toolbar=yes,scrollbars=yes,resizable=yes,top=100,left=200,width=600,height=600')</script>");
    }
    //寄送郵件 send mail fun
    protected void Button2_Click(object sender, EventArgs e)
    {
        //先記錄user維護的內容
        CommentRecord();
        //組目前的gd2的條件字串
        ContentQry();
        //string connectionString = "Data Source=GHICNEOR.WORLD;User ID=sd;PASSWORD=sd";
        string sMail_content=null;
        string connectionString = "Data Source=EC2;Initial Catalog=EC2;User ID=EC2;PASSWORD=gmtc";

        String sqlstr = @"select TOSHIBA_PROJECT,GMTC_PO,ITEM,WIP_LOT,FUJI_PO,TSB_PO,ROW_NUM from dbo.OV_ORDER_TRACE_JAP100 where cus_id ='" + Session["user"].ToString().Substring(0, 6) + @"'" + Global.sqlstring2 + Global.sWhere2+"order by row_num";
        using (SqlConnection conns = new SqlConnection(connectionString))
        {
            conns.Open();


            //設定查詢
            SqlCommand cmds = conns.CreateCommand();
            //cmds.CommandText = "SELECT COUNT(*) AS DCOUNT FROM OV_CERT_DOWNLOAD WHERE CUS_ID = @CUS_ID " + Global.sqlstring2;

            cmds.CommandText = sqlstr;   
            //OracleParameter param = new OracleParameter(":CUS_ID", OracleDbType.Varchar2);
            //param.Value = Session["user"];
            //cmds.Parameters.Add(param);
            int index=0;
            int count = 0;
            using (SqlDataReader ora = cmds.ExecuteReader())
            {

                while (ora.Read())
                {
                    index = Global.lb_rnum_gd2.IndexOf(ora["ROW_NUM"].ToString());
                    //Label6.Text = odr["total"].ToString();
                    //sMail_content = sMail_content + "\n " + ora["TOSHIBA_PROJECT"].ToString().PadRight(25) + "\t" + ora["GMTC_PO"].ToString().PadRight(10) + "\t" + ora["ITEM"].ToString().PadRight(4) + "\t" + ora["WIP_LOT"].ToString().PadRight(10) + "\t" + Global.txComment_gd2[index];
                    //sMail_content = sMail_content + "<tr><td>" + ora["TOSHIBA_PROJECT"].ToString().PadRight(25) + "</td><td>" + ora["GMTC_PO"].ToString().PadRight(10) + "</td><td>" + ora["ITEM"].ToString().PadRight(4) + "</td><td>" + ora["WIP_LOT"].ToString().PadRight(10) + "</td><td>"
                    //                + ora["FUJI_PO"].ToString().PadRight(10) + "</td><td>" + ora["TSB_PO"].ToString().PadRight(10) + "</td><td>" + Global.txComment_gd2[index] + "</td></tr></table>";
                    if (index!=-1)
                    {
                        //判斷是否為陣列最後一筆
                        //string bound = ora["ROW_NUM"].ToString();
                        if (count == Global.cb_list.Count)
                        {
                            index = Global.lb_rnum_gd2.IndexOf(ora["ROW_NUM"].ToString());
                            sMail_content = sMail_content + "<tr><td>" + ora["TOSHIBA_PROJECT"].ToString().PadRight(25) + "</td><td>" + ora["GMTC_PO"].ToString().PadRight(10) + "</td><td>" + ora["ITEM"].ToString().PadRight(4) + "</td><td>" + ora["WIP_LOT"].ToString().PadRight(10) + "</td><td>"
                            + ora["FUJI_PO"].ToString().PadRight(10) + "</td><td>" + ora["TSB_PO"].ToString().PadRight(10) + "</td><td>" + Global.txComment_gd2[index] + "</td></tr></table>";
                        }
                        else
                        {
                            index = Global.lb_rnum_gd2.IndexOf(ora["ROW_NUM"].ToString());
                            sMail_content = sMail_content + "<tr><td>" + ora["TOSHIBA_PROJECT"].ToString().PadRight(25) + "</td><td>" + ora["GMTC_PO"].ToString().PadRight(10) + "</td><td>" + ora["ITEM"].ToString().PadRight(4) + "</td><td>" + ora["WIP_LOT"].ToString().PadRight(10) + "</td><td>"
                            + ora["FUJI_PO"].ToString().PadRight(10) + "</td><td>" + ora["TSB_PO"].ToString().PadRight(10) + "</td><td>" + Global.txComment_gd2[index] + "</td></tr>";

                        }
                    }
                    //找不到該值
                    else if (index == -1)
                    {
                       
                        //string blank = ""; 
                                                //判斷是否為陣列最後一筆
                        //string bound = ora["ROW_NUM"].ToString();
                        if (count == Global.cb_list.Count)
                        {
                            sMail_content = sMail_content + "<tr><td>" + ora["TOSHIBA_PROJECT"].ToString().PadRight(25) + "</td><td>" + ora["GMTC_PO"].ToString().PadRight(10) + "</td><td>" + ora["ITEM"].ToString().PadRight(4) + "</td><td>" + ora["WIP_LOT"].ToString().PadRight(10) + "</td><td>"
                            + ora["FUJI_PO"].ToString().PadRight(10) + "</td><td>" + ora["TSB_PO"].ToString().PadRight(10) + "</td> <td> </td></tr></table>";
                        }
                        else
                        {
                            sMail_content = sMail_content + "<tr><td>" + ora["TOSHIBA_PROJECT"].ToString().PadRight(25) + "</td><td>" + ora["GMTC_PO"].ToString().PadRight(10) + "</td><td>" + ora["ITEM"].ToString().PadRight(4) + "</td><td>" + ora["WIP_LOT"].ToString().PadRight(10) + "</td><td>"
                            + ora["FUJI_PO"].ToString().PadRight(10) + "</td><td>" + ora["TSB_PO"].ToString().PadRight(10) + "</td> <td> </td></tr>";

                        }
                        
                    }
                    count++;
                }
                //count++;
            }
            
        }
        List<string> mailList = mail_get(); 
        //先將收件人組合起來
        string sReceiver = "";
        string sMail_header = "<table border=1> <tr><th>" + "TOSHIBA_PROJECT".PadRight(25) + "</th><th>" + "GMTC_PO".PadRight(10) + "</th><th>" + "ITEM".PadRight(6) + "</th><th>" + "WIP_LOT".PadRight(10) + "</th><th>" + "FUJI_PO".PadRight(10) + "</th><th>" + "TSB_PO".PadRight(14) + "</th><th>" + "COMMENT".PadRight(14) + "</th></tr>";
        //string sMail_header = "<table border=1> <tr><th>TOSHIBA_PROJECT</th><th>GMTC_PO</th><th>ITEM</th><th>WIP_LOT</th><th>FUJI_PO</th><th>TSB_PO</th><th>COMMENT</th></tr>";
        sMail_content = sMail_header + sMail_content;
        Global.sMail_Record = mailList[1];
        sReceiver = mailList[4];
        /* 註解list內容
        mailList[0] = sQUSTION_TYPE;//問題類型
        mailList[1] = sMAIL_TITLE;//郵件標題
        mailList[2] = sMAIL_SEND_CC;//副本收件人姓名
        mailList[3] = sMAIL_SEND_GMTC;//密件副本收件人email
		mailList[4] = sEMAIL;//收件人email
		mailList[5] = sEMAIL_CC//副本收件人email
        */
        //設定SMTP
        //德在大哥建議用.1.4去寄 by YenChang 07.15.2016
        SmtpClient MySmtp = new SmtpClient("192.168.1.4", 25);
        MailMessage mail = new MailMessage();
        //sMail_content = "Dear \r\n  Order Track list. Please remind customer send P/O for you. Thank you." + "\r\n\n\n " + TextBox7.Text + "\n\n\n" + sMail_content;
        //sMail_content = "Dear Sirs, \r\n\n  Please be noted customers have following questions thru order track system.\n Your early reply will be highly appreciated." + "\r\n\n\n " + sMail_content;
        sMail_content = "Dear Sirs, <br><br>  Please be noted customers have following questions thru order track system.<br> Your early reply will be highly appreciated.<br><br><br> " + sMail_content;
        //mail.Subject = DropDownList2.Text;
        mail.Subject = mailList[1];
        mail.IsBodyHtml = true;
        mail.From = new MailAddress("admin@gmtc.com.tw"); //設定寄件者;
        mail.To.Add(sReceiver);//這是收件者
        mail.Bcc.Add(mailList[3]); //這是密件副本收件者
        mail.CC.Add(mailList[5]);//這是副本收件者
        mail.Body = sMail_content;//內文
        MySmtp.Send(mail);
        //發送Email("寄信者","收信者","主題","內容")
        //MySmtp.Send("admin@gmtc.com.tw", "alex.yang@gmtc.com.tw", "OTD order test", @"Dear 您好:" + "\r\n" + Session["user"] + " Order OTD List. Please remind customer send P/O for you. Thank you." + "\r\n " + mail_content);
        //MySmtp.Send("admin@gmtc.com.tw", sReceiver, "Ord Track Test", "Dear \r\n  Order Track list. Please remind customer send P/O for you. Thank you." + "\r\n\n\n " + sMail_content);
        //Response.Write("<script language=javascript>alert('Send Mail Successed!!');location.href='Ordtrack.aspx';</script>");
        Response.Write("<script language=javascript>alert('Your mail has been sent successfully.');</script>");
        GridView2.Visible = true;
        String sqlstr2 = @"select TOSHIBA_PROJECT,GMTC_PO,ITEM,WIP_LOT,FUJI_PO,TSB_PO,ROW_NUM from dbo.OV_ORDER_TRACE_JAP100 where cus_id ='" + Session["user"].ToString().Substring(0, 6) + @"'" + Global.sqlstring2 + Global.sWhere2;
        SqlDataSource2.SelectCommand = sqlstr2;
        CommentRecord();
        Mail_Record();
        GridView2.DataBind();
    }
    protected void GridView1_DataBound(object sender, EventArgs e)
    {

        if (Global.sDelay == "true" && Global.sDelay != null)
        {
            Global.sWhere_Delay = "and DELAY>0";
        }
        else
        {
            Global.sWhere_Delay = "";
        }



        try
        {

            //資料庫連線設定
            string ConnectionStrings = "Data Source=EC2;user id=EC2;password=gmtc";
            using (SqlConnection conns = new SqlConnection(ConnectionStrings))
            {
                conns.Open();


                //設定查詢
                SqlCommand cmds = conns.CreateCommand();
                String ord_track = @"select count(*) AS DCOUNT from dbo.OV_ORDER_TRACE_JAP100 where cus_id ='" + Session["user"].ToString().Substring(0, 6) + @"'" + Global.sqlstring2 + Global.sWhere_Delay + " AND ROW_NUM <> '0'";
                cmds.CommandText = ord_track;
                //cmds.CommandText = "SELECT COUNT(*) AS DCOUNT FROM OV_CERT_DOWNLOAD WHERE CUS_ID = @CUS_ID " + Global.sqlstring;

                //SqlParameter param = new SqlParameter("@CUS_ID", SqlDbType.VarChar);
                //param.Value = Session["user"];
                //cmds.Parameters.Add(param);

                using (SqlDataReader odr = cmds.ExecuteReader())
                {
                    while (odr.Read())
                    {
                        Label11.Text = odr["DCOUNT"].ToString();
                    }
                }


            }
            if (Label11.Text != "0")
            {
                //給予原本搜尋條件
                //CheckBox ChkBoxHeader = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");
                ((TextBox)GridView1.HeaderRow.FindControl("TextBox8")).Text = HiddenField1.Value;
                ((TextBox)GridView1.HeaderRow.FindControl("TextBox9")).Text = HiddenField2.Value;
                ((TextBox)GridView1.HeaderRow.FindControl("TextBox10")).Text = HiddenField3.Value;
                ((TextBox)GridView1.HeaderRow.FindControl("TextBox11")).Text = HiddenField5.Value;
                ((TextBox)GridView1.HeaderRow.FindControl("TextBox12")).Text = HiddenField6.Value;
                ((TextBox)GridView1.HeaderRow.FindControl("TextBox13")).Text = HiddenField7.Value;
                ((DropDownList)GridView1.HeaderRow.FindControl("DropDownList1")).SelectedValue = HiddenField8.Value;
                ((TextBox)GridView1.HeaderRow.FindControl("TextBox67")).Text =  Session["TSB_NEED_DATE_1"].ToString();
                ((TextBox)GridView1.HeaderRow.FindControl("TextBox68")).Text = Session["TSB_NEED_DATE_2"].ToString();
                /*
                if (HiddenField4.Value == "true")
                {
                    //CheckBox ChkBoxHeader = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");
                    ChkBoxHeader.Checked = true;
                }
                else
                {
                    //GridView1.HeaderRow.FindControl("CheckBox2").EnableViewState = false;
                    ChkBoxHeader.Checked = false;
                }
                */
            }
        }
        catch
        {
            //Session.Clear();
            //Response.Write("<script language=javascript>alert('Please Login In!');location.href='Login.aspx';</script>");

        }
    }
    protected void CheckBox2_CheckedChanged(object sender, EventArgs e)
    {

        CheckBox ChkBoxHeader = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");
        bool t;
        CheckedAllBox();
        /*
        if (ChkBoxHeader.Checked == true)
        {
            HiddenField4.Value = "true";
            t = true;
            CheckedAllBox();
        }
        else
        {
            HiddenField4.Value = "false";
            t = false;
            CheckedAllBox();
        }
        //CheckedAllBox();
        ContentQry();
        */
    }
    //檢查哪些CHECKBOX被勾選
    private void CheckedBox()
    {
        //ref site http://www.dotnetgallery.com/kb/resource18-Gridview-header-checkbox-select-and-deselect-all-rows-using-client-side-Jav.aspx.aspx
        
        CheckBox ChkBoxHeader = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");
        
        int d = GridView1.PageCount;
        //int grid_page_count = 0;
        // BY YENCHANG 單選使用
        bool[] values = new bool[GridView1.PageSize];
        for (int i = 0; i < GridView1.Rows.Count; i++)
        {
            CheckBox ChkBoxRows = (CheckBox)GridView1.Rows[i].FindControl("CheckBox1");
            //CheckBox ChkBoxRows = (CheckBox)row.FindControl("CheckBox1");
            /*
            if (ChkBoxHeader.Checked == true)
            {
                ChkBoxRows.Checked = true;
                values[i] = ChkBoxRows.Checked;
                HiddenField4.Value = "true";
            }
            else
            {*/
            if (ChkBoxHeader.Checked == true)
            {
                values[i] = ChkBoxRows.Checked;
            }
            // test
            else
            {
                values[i] = false;
            }
            //HiddenField4.Value = "false";
            //}
        }

        
        //解決全暫存變數問題
        Session["page" + GridView1.PageIndex] = values;
    }
    //gridview checkbox均不選
    private void UnCheckedBox()
    {
        //ref site http://www.dotnetgallery.com/kb/resource18-Gridview-header-checkbox-select-and-deselect-all-rows-using-client-side-Jav.aspx.aspx

        CheckBox chb;
        CheckBox ChkBoxHeader = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");
        bool[] values = (bool[])Session["page" + GridView1.PageIndex];

        int d = GridView1.PageCount;
        //int grid_page_count = 0;
        // BY YENCHANG 單選使用
        //bool[] values = (bool[])Session["page" + GridView1.PageIndex];
        for (int i = 0; i < GridView1.Rows.Count; i++)
        {
            CheckBox ChkBoxRows = (CheckBox)GridView1.Rows[i].FindControl("CheckBox1");
            ChkBoxRows.Checked = false;
            values[i] = ChkBoxRows.Checked;

        }
        Session["page" + GridView1.PageIndex] = values;
        
    }

    //全選,使GridView所有checkbox1均為true或為false
    //private void CheckedAllBox(bool checkboxed)
    private void CheckedAllBox()
    {
        //ref site http://www.dotnetgallery.com/kb/resource18-Gridview-header-checkbox-select-and-deselect-all-rows-using-client-side-Jav.aspx.aspx
        if (Label11.Text == "0")
        {
        }
        else
        {
            
            CheckBox ChkBoxHeader = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");
            int d = GridView1.PageCount;
            //int grid_page_count = 0;
            bool[] values = new bool[GridView1.PageSize];
            // BY YENCHANG 全選使用

            for (int i = 0; i < GridView1.Rows.Count; i++)
            {
                CheckBox ChkBoxRows = (CheckBox)GridView1.Rows[i].FindControl("CheckBox1");
                //CheckBox ChkBoxRows = (CheckBox)row.FindControl("CheckBox1");

                if (CheckBox4.Checked == true)
                {
                    ChkBoxRows.Checked = true;
                    //HiddenField4.Value = "true";
                    values[i] = ChkBoxRows.Checked;
                    Button2.Visible = true;
                    Label6.Visible = true;
                    DropDownList2.Visible = true;
                    GridView2.Visible = true;
                    //CheckBoxRecord();
                }
                else if (CheckBox5.Checked == true)
                {
                    //HiddenField4.Value = "false";
                    ChkBoxRows.Checked = false;
                    values[i] = false;
                    Button2.Visible = false;
                    Label6.Visible = false;
                    DropDownList2.Visible = false;
                    GridView2.Visible = false;
                }

            //Session["page" + GridView1.PageIndex] = values;
        }
            
        }
        
    }

    //Remove 去除勾選
    //REMOVE GRIDVIEW2
    protected void LinkButton1_Click(object sender, EventArgs e)
    {
    
        #region 記錄哪列改變
        LinkButton pb = (LinkButton)sender;
        GridViewRow row = (GridViewRow)pb.NamingContainer;
        #endregion
        Label lbrov = (row.Cells[5].FindControl("Label1") as Label);
        //Global.sWhere = null;
        Global.cb_list.Remove(lbrov.Text.ToString());
        List<String> gv2ronum_values = new List<String>();
        List<String> gv2comment = new List<String>();
        Label lb_rnum_gd2;
        TextBox tbComment_gd2;
        for (int i = 0; i < GridView2.Rows.Count; i++)
        {
            #region 記錄哪列改變
            tbComment_gd2 = (TextBox)GridView2.Rows[i].FindControl("TextBox44");
            
            GridViewRow row2 = (GridViewRow)tbComment_gd2.NamingContainer;
            lb_rnum_gd2 = row2.Cells[5].FindControl("Label1") as Label;
            tbComment_gd2 = row2.Cells[8].FindControl("TextBox44") as TextBox;
            #endregion
            //if (tbComment_gd2.ToString() != null)
            if (lb_rnum_gd2.Text != lbrov.Text && tbComment_gd2.Text!="")
            {
                //#region 記錄哪列改變
                //GridViewRow row = (GridViewRow)tbComment_gd2.NamingContainer;
                //#endregion
                gv2ronum_values.Add(lb_rnum_gd2.Text);
                gv2comment.Add(tbComment_gd2.Text);
               //Global.txComment_gd2.Add(tbComment_gd2.Text);

            }
            else 
            {


            }
            //GridView2.AllowPaging = true;
        }
        Session["gv2ronum"] = gv2ronum_values;
        Session["gv2comment"] = gv2comment;
        ContentQry();
        //CommentRecord();
        
        for (int j = 0; j < GridView1.PageCount; j++)
        {
            for (int k = 0; k < GridView1.Rows.Count; k++)
            {
                CheckBox chb = (CheckBox)GridView1.Rows[k].FindControl("CheckBox1");
                GridViewRow row3 = (GridViewRow)chb.NamingContainer;
                Label lb_rnum_gv1 = row3.Cells[24].FindControl("Label7") as Label;
                if (lb_rnum_gv1.Text == lbrov.Text)
                {
                    chb.Checked = false;
                }
            }
        }

    }
    //checkbox 控制gridview2 bind
    private void ContentQry()
    {
        Global.sWhere2 = null;
        if (Global.cb_list.Count != 0)
        {
            for (int i = 0; i < Global.cb_list.Count; i++)
            {
                if (Global.sWhere2 == null)
                {
                    Global.sWhere2 = " AND ( ROW_NUM = " + Global.cb_list[i] + ")";
                }

                else
                {
                    Global.sWhere2 = Global.sWhere2.Substring(0, Global.sWhere2.Length - 1) + " or ROW_NUM = " + Global.cb_list[i] + ")";
                }
            }
        }
        else
        {
            Global.sWhere2 = " AND ( ROW_NUM =-1)";
        }

        //Global.cb_list
        //try
        //{


        //String sqlstr = @"select aa.PO_NO,aa.ORDER_NO,aa.ORDER_SEQ,aa.CERT_PO,aa.WIP_LOT from (select PO_NO,ORDER_NO,ORDER_SEQ,CERT_PO,WIP_LOT,row_number() over (order by PO_NO asc) as [Row_Number] from dbo.OV_ORDER_TRACE_MAIN where cus_id ='JAP100') aa " + Global.sqlstring2 + Global.sWhere;
        String sqlstr = @"select TOSHIBA_PROJECT,GMTC_PO,ITEM,WIP_LOT,ROW_NUM,FUJI_PO,TSB_PO  from dbo.OV_ORDER_TRACE_JAP100 where cus_id ='" + Session["user"].ToString().Substring(0, 6) + @"'" + Global.sqlstring2 + Global.sWhere2;
        SqlDataSource2.SelectCommand = sqlstr;
        CommentRecord();
        GridView2.DataBind();
        //GridView2.AllowPaging = true;
        GridView2.Visible = true;
    }
    protected void GridView2_PageIndexChanging(object sender, GridViewPageEventArgs e)
    {
        GridView2.PageIndex = e.NewPageIndex;
        ContentQry();
        //CommentRecord();
        
    }
    protected void Button3_Click(object sender, EventArgs e)
    {
        this.Response.Write("<script language=javascript>window.open('Detail_Track.aspx','Detail','toolbar=yes,scrollbars=yes,resizable=yes,top=100,left=200,width=600,height=600')</script>");
    }

    protected void LinkButton2_Click2(object sender, EventArgs e)
    {
        #region 記錄哪列改變
        LinkButton pb = (LinkButton)sender;
        GridViewRow row = (GridViewRow)pb.NamingContainer;
        #endregion
        Label pblabel = (row.Cells[26].FindControl("Label1") as Label);
        Session["wip_lot"] = pblabel.Text;
        this.Response.Write("<script language=javascript>window.open('Detail_Track.aspx','Detail','toolbar=yes,scrollbars=yes,resizable=yes,top=100,left=200,width=600,height=600')</script>");
    }
    protected void TextBox9_TextChanged(object sender, EventArgs e)
    {

    }
    protected void Button4_Click(object sender, EventArgs e)
    {

    }





    /*
    protected void CheckBox1_CheckedChanged(object sender, EventArgs e)
    {

    }
    */
    protected void TextBox12_TextChanged(object sender, EventArgs e)
    {
        
    }

    protected void Button5_Click(object sender, EventArgs e)
    {
        clear_condition();
        /*
        //清空所有條件
        Global.sqlstring2 = null;

        HiddenField1.Value = "";
        HiddenField2.Value = "";
        HiddenField3.Value = "";
        //HiddenField4.Value = "";
        HiddenField5.Value = "";
        HiddenField6.Value = "";
        HiddenField7.Value = "";
        HiddenField8.Value = "";
        HiddenField4.Value = "false";
        Global.cb_list.Clear();
        //重新搜尋所有條件
        String ord_track = @"select TOSHIBA_PROJECT,GMTC_PO,ITEM,FUJI_PO,GRADE,SIZE_,TSB_PO,EX_WORKS,ORD_QTY,SHIPEED_QTY,IN_STOCK_QTY,WIP_QTY,IN_STOCK_PCS,PLAN_TO_SHIP_FROM_GMTC,CSD,
        ETD,ETA,TSB_NEED_DATE,SHIPPING_STATUS,WIP,WEIGHT_TOLERANCE,DELAY,NOTE,ROW_NUM,CUS_ID,WIP_LOT from dbo.OV_ORDER_TRACE_JAP100 where cus_id = '" + Session["user"] + @"'" + Global.sqlstring2;
        CheckedBox();
        String sqlstr = @"select TOSHIBA_PROJECT,GMTC_PO,ITEM,WIP_LOT,ROW_NUM  from dbo.OV_ORDER_TRACE_JAP100 where cus_id ='-1'";
        SqlDataSource2.SelectCommand = sqlstr;
        GridView2.DataBind();
        SqlDataSource1.SelectCommand = ord_track;
        SqlDataSource1.SelectParameters.Clear();
        */
        /*
        if (HiddenField4.Value == "false")
        {
            CheckBox ChkBox = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox1");
            int d = GridView1.PageCount;
            bool[] values2 = new bool[GridView1.PageSize];

            for (int i = 0; i < GridView1.Rows.Count; i++)
            {
                    ChkBox.Checked = false;
                    values2[i] = ChkBox.Checked;
                    HiddenField4.Value = "false";

            }
            Session["page" + GridView1.PageIndex] = values2;
        }
         */
        //GridView1.DataBind();
        //Session["page" + GridView1.PageIndex] = values2;
    }
    // 新增 clear fun
    private void clear_condition()
    {
        //清空所有條件
        Global.sqlstring2 = null;
        for (int i = 0; i < GridView1.PageCount; i++)
        {
            Session["page" + i] = null;//清除pageSESSION
        }
        Global.sWhere_Delay = null;
        Session["gv2ronum"] = null;
        Session["gv2comment"] = null;
        Session["TSB_NEED_DATE_1"] = null;
        Session["TSB_NEED_DATE_2"] = null;
        Global.sDelay="false";
        HiddenField1.Value = "";
        HiddenField2.Value = "";
        HiddenField3.Value = "";
        //HiddenField4.Value = "";
        HiddenField5.Value = "";
        HiddenField6.Value = "";
        HiddenField7.Value = "";
        HiddenField8.Value = "";
        HiddenField4.Value = "false";
        Global.cb_list.Clear();
        Global.lb_rnum_gd2.Clear();
        Global.txComment_gd2.Clear();
        //重新搜尋所有條件
        String ord_track = @"select TOSHIBA_PROJECT,GMTC_PO,ITEM,FUJI_PO,GRADE,SIZE_,TSB_PO,EX_WORKS,ORD_QTY,SHIPEED_QTY,IN_STOCK_QTY,WIP_QTY,IN_STOCK_PCS,PLAN_TO_SHIP_FROM_GMTC,CSD,
        ETD,ETA,TSB_NEED_DATE,SHIPPING_STATUS,WIP,WEIGHT_TOLERANCE,DELAY,NOTE,ROW_NUM,CUS_ID,WIP_LOT from dbo.OV_ORDER_TRACE_JAP100 where cus_id = '" + Session["user"].ToString().Substring(0, 6) + @"'" + Global.sqlstring2 + " AND ROW_NUM <> '0' ORDER BY EX_WORKS";
        Label6.Visible = false;
        DropDownList2.Visible = false;
        Button2.Visible = false;
        GridView2.Visible = false;
        //(CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");
        //CheckBox ChkBoxHeader = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");
        if (Label11.Text == "0")
        {
            //CheckBox ChkBoxHeader = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");

        }
        else
        {
            CheckBox ChkBoxHeader = (CheckBox)GridView1.HeaderRow.FindControl("CheckBox2");
            ChkBoxHeader.Checked = false;
        }
        CheckedAllBox();
        String sqlstr = @"select TOSHIBA_PROJECT,GMTC_PO,ITEM,WIP_LOT,ROW_NUM,FUJI_PO,TSB_PO from dbo.OV_ORDER_TRACE_JAP100 where cus_id ='-1'";
        SqlDataSource2.SelectCommand = sqlstr;
        GridView2.DataBind();
        SqlDataSource1.SelectCommand = ord_track;
        SqlDataSource1.SelectParameters.Clear();
        CommentRecord();
        CheckBox4.Checked = false;
        CheckBox5.Checked = false;
    }
    protected void Button6_Click(object sender, EventArgs e)
    {

        String ord_track = @"select TOSHIBA_PROJECT,GMTC_PO,ITEM,FUJI_PO,GRADE,SIZE_,TSB_PO,CONVERT(VARCHAR(10) ,EX_WORKS, 101 ) EX_WORKS,ORD_QTY,SHIPEED_QTY SHIPPED_QTY ,IN_STOCK_QTY,WIP_QTY,IN_STOCK_PCS,CONVERT(VARCHAR(10) , PLAN_TO_SHIP_FROM_GMTC, 101 ) PLAN_TO_SHIP_FROM_GMTC,CONVERT(VARCHAR(10) , CSD, 101 ) CSD,
            CONVERT(VARCHAR(10) , ETD, 101 ) ETD ,CONVERT(VARCHAR(10) , ETA, 101 ) ETA ,CONVERT(VARCHAR(10) , TSB_NEED_DATE, 101 ) TSB_NEED_DATE,SHIPPING_STATUS,WIP,WEIGHT_TOLERANCE,DELAY,NOTE,FIC_COMMENT from dbo.OV_ORDER_TRACE_JAP100 where CUS_ID = '" + Session["user"].ToString().Substring(0, 6) + @"'" + Global.sqlstring2 + Global.sWhere_Delay + " AND ROW_NUM <> '0'";
        String connecting = "Data Source=EC2;Initial Catalog=EC2;User ID=EC2;PASSWORD=gmtc";
        ToExcel(ord_track);
        //UPLOAD XLS
        /*
        GridView3.Visible = true;
        SqlDataSource3.SelectCommand = ord_track;
        GridView3.DataBind();
        string sPath = ToExcel_New(ord_track);
        DownlodFile(sPath);
        
        SqlDataSource3.SelectCommand = ord_track;
        SqlDataSource3.SelectParameters.Clear();
        GridView3.DataBind();
        GridView3.AllowPaging = false;
        GridView3.Visible = true;
        GridView3.DataSource = (DataTable)ViewState["ExportTable"];
        string excelFileName = "Result.xls";
        Response.Clear();

        Context.Response.AddHeader("content-disposition", "attachment;filename=" +
        Server.UrlEncode(excelFileName));
        Response.Cache.SetCacheability(HttpCacheability.NoCache);
        Context.Response.ContentType = "application/vnd.xls";
        Response.Charset = "utf-8"; // "big5";  
        System.IO.StringWriter tw = new System.IO.StringWriter();
        HtmlTextWriter hw = new HtmlTextWriter(tw);
        this.GridView3.RenderControl(hw);

        Context.Response.Write(tw.ToString().Replace("<div>", "").Replace("</div>", ""));
        Context.Response.End();

        GridView3.AllowPaging = true;
        GridView3.Visible = false;
        GridView3.DataBind();
        */
       
    }
    //參考網址:http://www.aspsnippets.com/Articles/Export-DataSet-or-DataTable-to-Word-Excel-PDF-and-CSV-Formats.aspx
    public override void VerifyRenderingInServerForm(Control control)
    {
        /* Confirms that an HtmlForm control is rendered for the specified ASP.NET
           server control at run time. */
    }
    
    private void ToExcel(string sqlstr)
    {
        //Get the data from database into datatable

        SqlCommand cmd = new SqlCommand(sqlstr);
        DataTable dt = GetData(cmd);

        //Create a dummy GridView
        GridView GridView1 = new GridView();
        GridView1.AllowPaging = false;
        GridView1.DataSource = dt;
        GridView1.DataBind();
        GridView1.Font.Name = "Arial";
        Response.Clear();
        Response.Buffer = true;
        Response.AddHeader("content-disposition",
         "attachment;filename=DataTable.xls");
        Response.Charset = "";
        Response.ContentType = "application/vnd.ms-excel";
        StringWriter sw = new StringWriter();
        HtmlTextWriter hw = new HtmlTextWriter(sw);

        for (int i = 0; i < GridView1.Rows.Count; i++)
        {
            //Apply text style to each Row
            GridView1.Rows[i].Attributes.Add("class", "textmode");
        }
        GridView1.RenderControl(hw);

        //style to format numbers to string
        //string style = @"<style> .textmode { mso-number-format:\@; } </style>";
        Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>");
        //Response.Write(style);
        Response.Output.Write(sw.ToString());
        Response.Flush();
        Response.End();
    }

    
    private string  ToExcel_New(string sqlstr)
    {
        string excelFilename = "ExportExcel";
        Microsoft.Office.Interop.Excel.Application objexcelapp = new Microsoft.Office.Interop.Excel.Application();
        //objexcelapp.Application.Workbooks.Add(Type.Missing);
        objexcelapp.Application.Workbooks.Add(true);
        //objexcelapp.Visible = true;
        objexcelapp.Visible = false;
        objexcelapp.Columns.ColumnWidth = 25;
        string path = Server.MapPath("exportedfiles\\");

        if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
        {
            Directory.CreateDirectory(path);
        }

        File.Delete(path + "ExportExcel.xls"); // DELETE THE FILE BEFORE CREATING A NEW ONE.
        //SqlCommand cmd = new SqlCommand(sqlstr);
        //DataTable dt = GetData(cmd);

        //Create a dummy GridView
        /*
        GridView GridView1 = new GridView();
        GridView1.AllowPaging = false;
        GridView1.AutoGenerateColumns = true;
        GridView1.DataSource = dt;
        GridView1.DataBind();
        Response.Clear();
        Response.Buffer = true;
        Response.AddHeader("content-disposition",
         "attachment;filename=DataTable.xls");
        Response.Charset = "";
        Response.ContentType = "application/vnd.ms-excel";
        StringWriter sw = new StringWriter();
        HtmlTextWriter hw = new HtmlTextWriter(sw);

        for (int i = 0; i < GridView1.Rows.Count; i++)
        {
            //Apply text style to each Row
            GridView1.Rows[i].Attributes.Add("class", "textmode");
        }
        GridView1.RenderControl(hw);
        */
        //int NoOfColumns = GridView1.Rows[0].Cells.Count;
        for (int i = 1; i < GridView3.Columns.Count + 1; i++)
        {
            objexcelapp.Cells[1, i] = GridView3.Columns[i - 1].HeaderText;
            Microsoft.Office.Interop.Excel.Range headRange = objexcelapp.Cells[1, i] as Microsoft.Office.Interop.Excel.Range;
            headRange.EntireColumn.AutoFit();//自動調整列寬
            //自動篩選
            //headRange.EntireColumn.AutoFilter();
            headRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;//設置邊框
            //自動篩選
            //headRange.EntireColumn.AutoFilter();
        }
        /*For storing Each row and column value to excel sheet*/
    
        for (int i = 0; i < GridView3.Rows.Count; i++)
        {
            for (int j = 0; j < GridView3.Columns.Count; j++)
            {
                //if (GridView1.Rows[i].Cells[j].Text.ToString() != null)
                //{ 
                string sTemp = null;
                if (GridView3.Rows[i].Cells[j].Text == "&nbsp;")
                {
                    sTemp = GridView3.Rows[i].Cells[j].Text.Replace("&nbsp;", " ");
                }
                else
                {
                    sTemp = GridView3.Rows[i].Cells[j].Text.Replace("&#39;", "");
                }
                //sTemp.Replace("&nbsp;", "");
                objexcelapp.Cells[i + 2, j + 1] = sTemp ;
                Microsoft.Office.Interop.Excel.Range contentRange = objexcelapp.Cells[i + 2, j + 1] as Microsoft.Office.Interop.Excel.Range;
                contentRange.EntireColumn.AutoFit();//自動調整列寬
                contentRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;//設置邊框
                //}
            }
         }
        //MessageBox.Show("Your excel file exported successfully at D:\\權限稽核\\2016.06\\" + Global.excelFilename + ".xls");
        //objexcelapp.ActiveWorkbook.SaveCopyAs("D:\\權限稽核\\2016.06\\" + Global.excelFilename + ".xls");
        //objexcelapp.ActiveWorkbook.Saved = true;
        Microsoft.Office.Interop.Excel.Range iRng = objexcelapp.get_Range("A1", "A1");// From
        //自動欄寬
        //iRng.AutoFilter(1, Type.Missing, Microsoft.Office.Interop.Excel.XlAutoFilterOperator.xlAnd, Type.Missing, true);
        objexcelapp.DisplayAlerts = false;
        objexcelapp.AlertBeforeOverwriting = false;
        string sPath = path + excelFilename + ".xls";
        objexcelapp.ActiveWorkbook.SaveAs(sPath);

        //MessageBox.Show("D:\\" + excelFilename + ".xls");
        objexcelapp.ActiveWorkbook.Saved = true;
        // CLEAR. 變數需RELEASE 不然會導致資料錯誤
        objexcelapp.Workbooks.Close();
        objexcelapp.Quit();
        objexcelapp = null;
        GridView3.Visible = false;
        return sPath;

    }
    private void DownlodFile(string sPath)
    {
        try
        {
            //string sPath = Server.MapPath("exportedfiles\\");
            Response.AppendHeader("Content-Disposition", "attachment; filename=ExportExcel.xls");
            //Response.TransmitFile(path + "ExportExcel.xls");
            Response.ContentType = "application/download";
            Response.TransmitFile(sPath);
            Response.End();
            //objexcelapp.Quit();
        }
        catch (Exception ex) 
        {
           Response.Write(ex);
        }

    }

  

    private DataTable GetData(SqlCommand cmd)
    {
        DataTable dt = new DataTable();
        String strConnString = System.Configuration.ConfigurationManager.
             ConnectionStrings["EC2ConnectionString4"].ConnectionString;
        SqlConnection con = new SqlConnection(strConnString);
        SqlDataAdapter sda = new SqlDataAdapter();
        cmd.CommandType = CommandType.Text;
        cmd.Connection = con;
        try
        {
            con.Open();
            sda.SelectCommand = cmd;
            sda.Fill(dt);
            return dt;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
            con.Close();
            sda.Dispose();
            con.Dispose();
        }
    }
    
    protected void Button7_Click(object sender, EventArgs e)
    {
        
        String ord_track = @"select TOSHIBA_PROJECT,GMTC_PO,ITEM,FUJI_PO,GRADE,SIZE_,TSB_PO,EX_WORKS,ORD_QTY,SHIPEED_QTY,IN_STOCK_QTY,WIP_QTY,IN_STOCK_PCS,PLAN_TO_SHIP_FROM_GMTC,CSD,
            ETD,ETA,TSB_NEED_DATE,SHIPPING_STATUS,WIP,WEIGHT_TOLERANCE,DELAY,NOTE,ROW_NUM,CUS_ID,WIP_LOT from dbo.OV_ORDER_TRACE_JAP100 where CUS_ID = '" + Session["user"] + @"'" + Global.sqlstring2;
        ToWord(ord_track);
        
       

    }
    
    private void ToWord(string sqlstr)
    {

        SqlCommand cmd = new SqlCommand(sqlstr);
        DataTable dt = GetData(cmd);

        //Create a dummy GridView
        GridView GridView1 = new GridView();
        GridView1.AllowPaging = false;
        GridView1.DataSource = dt;
        GridView1.DataBind();

        Response.Clear();
        Response.Buffer = true;
        Response.AddHeader("content-disposition",
            "attachment;filename=DataTable.doc");
        Response.Charset = "";
        Response.ContentType = "application/vnd.ms-word ";
        StringWriter sw = new StringWriter();
        HtmlTextWriter hw = new HtmlTextWriter(sw);
        GridView1.RenderControl(hw);
        Response.Output.Write(sw.ToString());
        Response.Flush();
        Response.End();
    }
    
    protected void Button8_Click(object sender, EventArgs e)
    {
        
        String ord_track = @"select TOSHIBA_PROJECT,GMTC_PO,ITEM,FUJI_PO,GRADE,SIZE_,TSB_PO,EX_WORKS,ORD_QTY,SHIPEED_QTY,IN_STOCK_QTY,WIP_QTY,IN_STOCK_PCS,PLAN_TO_SHIP_FROM_GMTC,CSD,
            ETD,ETA,TSB_NEED_DATE,SHIPPING_STATUS,WIP,WEIGHT_TOLERANCE,DELAY,NOTE,ROW_NUM,CUS_ID,WIP_LOT from dbo.OV_ORDER_TRACE_JAP100 where CUS_ID = '" + Session["user"] + @"'" + Global.sqlstring2;
        ToPDF(ord_track);
       

    }
    
    private void ToPDF(string sqlstr)
    {
        //Get the data from database into datatable
        SqlCommand cmd = new SqlCommand(sqlstr);
        DataTable dt = GetData(cmd);

        //Create a dummy GridView
        GridView GridView1 = new GridView();
        GridView1.AllowPaging = false;
        GridView1.DataSource = dt;
        GridView1.DataBind();

        Response.ContentType = "application/pdf";
        Response.AddHeader("content-disposition",
            "attachment;filename=DataTable.pdf");
        Response.Cache.SetCacheability(HttpCacheability.NoCache);
        StringWriter sw = new StringWriter();
        HtmlTextWriter hw = new HtmlTextWriter(sw);
        GridView1.RenderControl(hw);
        StringReader sr = new StringReader(sw.ToString());
        Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 0f);
        HTMLWorker htmlparser = new HTMLWorker(pdfDoc);
        PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
        pdfDoc.Open();
        htmlparser.Parse(sr);
        pdfDoc.Close();
        Response.Write(pdfDoc);
        Response.End();
    }

    public void Query_Record()
    {
        //UN_Lock_LogTable();
        //資料庫連線設定
        string ConnectionStrings = "Data Source=EC2;user id=ec2;password=gmtc";
        string KIND_TYPE = "Query";
        using (SqlConnection connection = new SqlConnection(ConnectionStrings))
        {
            connection.Open();

            //設定交易
            SqlTransaction transaction = connection.BeginTransaction();

            //設定INSERT
            SqlCommand command = connection.CreateCommand();
            command.Transaction = transaction;
            //Session["cus_ip"] = Request.ServerVariables["REMOTE_ADDR"].ToString();
            try
            {
                //cmds1_Record.CommandText = @"INSERT INTO OV_TRACE_RECORD  (CUS_ID, PO_NO, ORDER_NO, ORDER_SEQ, WIP_LOT, IP_ADDRESS, IP_TIME, KIND_TYPE) VALUES ('" + Session["user"] + "', '', '', '', '', '" +Request.ServerVariables["REMOTE_ADDR"] + "', '" + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + "', 'Login')";
                command.CommandText = @"INSERT INTO OV_TRACE_RECORD  (CUS_ID, PO_NO, ORDER_NO, ORDER_SEQ, WIP_LOT, IP_ADDRESS, IP_TIME, KIND_TYPE,DESCR) 
                                                      VALUES ('" + Session["user"] + "', '" + Session["user"] + "', '" + Session["user"] + "', '" +
                                   Session["user"] + "', '" + Session["user"] + "', '" +
                                   Request.ServerVariables["REMOTE_ADDR"] + "', '" + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + "','" + KIND_TYPE + "','" + Global.sRecord + "')";

                command.ExecuteNonQuery();
                command.Parameters.Clear();
                transaction.Commit();
                connection.Close();
                
            }
            catch (Exception ex)
            {
                transaction.Rollback();
                connection.Close();
            }
            
        }
    }
    // YenChang 加入寄送郵件內容
    public void Mail_Record()
    {
        //UN_Lock_LogTable();
        //資料庫連線設定
        string ConnectionStrings = "Data Source=EC2;user id=ec2;password=gmtc";
        string KIND_TYPE = "Mail";
        using (SqlConnection connection = new SqlConnection(ConnectionStrings))
        {
            connection.Open();

            //設定交易
            SqlTransaction transaction = connection.BeginTransaction();

            //設定INSERT
            SqlCommand command = connection.CreateCommand();
            command.Transaction = transaction;
            //Session["cus_ip"] = Request.ServerVariables["REMOTE_ADDR"].ToString();
            try
            {
                //cmds1_Record.CommandText = @"INSERT INTO OV_TRACE_RECORD  (CUS_ID, PO_NO, ORDER_NO, ORDER_SEQ, WIP_LOT, IP_ADDRESS, IP_TIME, KIND_TYPE) VALUES ('" + Session["user"] + "', '', '', '', '', '" +Request.ServerVariables["REMOTE_ADDR"] + "', '" + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + "', 'Login')";
                command.CommandText = @"INSERT INTO OV_TRACE_RECORD  (CUS_ID, PO_NO, ORDER_NO, ORDER_SEQ, WIP_LOT, IP_ADDRESS, IP_TIME, KIND_TYPE,DESCR) 
                                                      VALUES ('" + Session["user"] + "', '" + Session["user"] + "', '" + Session["user"] + "', '" +
                                   Session["user"] + "', '" + Session["user"] + "', '" +
                                   Request.ServerVariables["REMOTE_ADDR"] + "', '" + DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss") + "','" + KIND_TYPE + "','" + Global.sMail_Record + "')";

                command.ExecuteNonQuery();
                command.Parameters.Clear();
                transaction.Commit();
                connection.Close();

            }
            catch (Exception ex)
            {
                transaction.Rollback();
                connection.Close();
            }

        }
    }

    
    private void CheckboxallRecord()
    {

        string connectionString = "Data Source=EC2;Initial Catalog=EC2;User ID=EC2;PASSWORD=gmtc";
        String sqlstr = @"select [ROW_NUM] from dbo.OV_ORDER_TRACE_JAP100 where cus_id ='" + Session["user"] + "'" + Global.sqlstring2 + " AND ROW_NUM <> '0'";
        Global.cb_list.Clear();
        if (CheckBox4.Checked == true)
        {
            using (SqlConnection conns = new SqlConnection(connectionString))
            {
                conns.Open();


                //設定查詢
                SqlCommand cmds = conns.CreateCommand();
                //cmds.CommandText = "SELECT COUNT(*) AS DCOUNT FROM OV_CERT_DOWNLOAD WHERE CUS_ID = @CUS_ID " + Global.sqlstring2;

                cmds.CommandText = sqlstr;
                //OracleParameter param = new OracleParameter(":CUS_ID", OracleDbType.Varchar2);
                //param.Value = Session["user"];
                //cmds.Parameters.Add(param);

                using (SqlDataReader ora = cmds.ExecuteReader())
                {

                    while (ora.Read())
                    {
                        //Label6.Text = odr["total"].ToString();
                        //sMail_content = sMail_content + "\n" + ora["TOSHIBA_PROJECT"].ToString().PadRight(25) + "\t" + ora["GMTC_PO"] + "\t" + ora["ITEM"] + "\t" + ora["WIP_LOT"];
                        Global.cb_list.Add(ora["ROW_NUM"].ToString());
                    }
                }
            }

        }
        else if(CheckBox5.Checked==true)
        {
            Global.cb_list.Clear();
        }
        /*
        using (SqlConnection conns = new SqlConnection(connectionString))
        {
            conns.Open();


            //設定查詢
            SqlCommand cmds = conns.CreateCommand();
            //cmds.CommandText = "SELECT COUNT(*) AS DCOUNT FROM OV_CERT_DOWNLOAD WHERE CUS_ID = @CUS_ID " + Global.sqlstring2;

            cmds.CommandText = sqlstr;
            //OracleParameter param = new OracleParameter(":CUS_ID", OracleDbType.Varchar2);
            //param.Value = Session["user"];
            //cmds.Parameters.Add(param);

            using (SqlDataReader ora = cmds.ExecuteReader())
            {

                while (ora.Read())
                {
                    //Label6.Text = odr["total"].ToString();
                    //sMail_content = sMail_content + "\n" + ora["TOSHIBA_PROJECT"].ToString().PadRight(25) + "\t" + ora["GMTC_PO"] + "\t" + ora["ITEM"] + "\t" + ora["WIP_LOT"];
                    Global.cb_list.Add(ora["ROW_NUM"].ToString());
                }
            }
        }
       
        */
    }
    // 寫入log寫法
    // 亮哥表示要秀出資料更新時間 by YenChang 06.15.2016
    // 寫入update log時間
    private void Update_Show(string cus_id)
    {
        //file path
        string path ="c:\\UpdateTime_"+cus_id+".txt";
        //無檔案則建立一個
        if (!File.Exists(path))
        {
            //新增檔案
            System.IO.StreamWriter newfile = new System.IO.StreamWriter(path);   //每次都重頭寫入
            newfile.Flush();
            newfile.Close();
        }
        //取得目前時間
        System.DateTime currentTime = new System.DateTime(); 
        currentTime = System.DateTime.Now;
        System.IO.FileStream fs = new System.IO.FileStream(path, FileMode.Open);
        //已串流方式讀取
        StreamReader sr = new StreamReader(fs);
        string past_time = sr.ReadToEnd();
        DateTime past_time_dt;
        if (String.IsNullOrEmpty(past_time))
        {
            past_time_dt = currentTime;
        }
        else
        {
            past_time_dt = Convert.ToDateTime(past_time);
        }
        TimeSpan days = new TimeSpan(currentTime.Ticks - past_time_dt.Ticks);
        fs.Flush();
        fs.Close();
        int int_days = Convert.ToInt16(days.Days.ToString());
        //亮哥表示一週更新一次
        if (int_days >= 7)
        {
            //寫入log
            System.IO.StreamWriter file = new System.IO.StreamWriter(path);   //每次都重頭寫入
            file.WriteLine(currentTime);
            Label20.Text = String.Format("{0:MM/dd/yyyy}", past_time_dt);
            file.Flush();
            file.Close();
        }
        else
        {
            Label20.Text = String.Format("{0:MM/dd/yyyy}", past_time_dt);
        }

    }
    // 亮哥要求增加06.15 by YenChang
    //增加 "Delay Order" 查詢按鈕, 條件   ->     select * from dbo.OV_ORDER_TRACE_JAP100 where DELAY>0
    protected void Button9_Click(object sender, EventArgs e)
    {
        //查詢時先清除條件
        /*
        Button button = FindControl<Button>(this.Page, "Button5") as Button;
        //Button button = (Button)this.FindControl("ContentPlaceHolder2$Button5");
        button = (Button)sender;
        button.Click += new EventHandler(Button5_Click);
        */
        clear_condition();
        Global.sDelay = "true";
        Global.sWhere_Delay = "and DELAY>0";
        string sql = "select * from dbo.OV_ORDER_TRACE_JAP100 where 1=1" + Global.sWhere_Delay + " ORDER BY EX_WORKS ";
        SqlDataSource1.SelectCommand = sql;
        GridView1.DataBind();
        
    }
    //練習抓page頁面的控制項
    //參考連結http://blog.miniasp.com/post/2007/11/04/ASPNET-FindControl-Tips-and-Hacks.aspx
    public T FindControl<T>(string id) where T : Control
    {
        return FindControl<T>(Page, id);
    }

    public static T FindControl<T>(Control startingControl, string id) where T : Control
    {
        // 取得 T 的預設值，通常是 null
        T found = default(T);

        int controlCount = startingControl.Controls.Count;

        if (controlCount > 0)
        {
            for (int i = 0; i < controlCount; i++)
            {
                Control activeControl = startingControl.Controls[i];
                if (activeControl is T)
                {
                    found = startingControl.Controls[i] as T;
                    if (string.Compare(id, found.ID, true) == 0) break;
                    else found = null;
                }
                else
                {
                    found = FindControl<T>(activeControl, id);
                    if (found != null) break;
                }
            }
        }
        return found;
    }
    private void SqlUpdate()
    {
        /*
        string sql = @"SELECT DATEADD(wk,DATEDIFF(wk,0,getdate()),-2) 'FIR' from dual where getdate()<>DATEADD(wk,DATEDIFF(wk,0,getdate()),-2)
                        union all
                      SELECT DATEADD(wk,DATEDIFF(wk,0,getdate()),+5) 'FIR' from dual where getdate()=DATEADD(wk,DATEDIFF(wk,0,getdate()),+5)";
        */
        string sql = @"select UID_TIME 'FIR' from dbo.OV_ORDER_TRACE_JAP100 where row_num=0";
        string connectionString = "Data Source=EC2;Initial Catalog=EC2;User ID=EC2;PASSWORD=gmtc";
        
        using (SqlConnection conns = new SqlConnection(connectionString))
        {
            conns.Open();


            //設定查詢
            SqlCommand cmds = conns.CreateCommand();
            //cmds.CommandText = "SELECT COUNT(*) AS DCOUNT FROM OV_CERT_DOWNLOAD WHERE CUS_ID = @CUS_ID " + Global.sqlstring2;

            cmds.CommandText = sql;
            //OracleParameter param = new OracleParameter(":CUS_ID", OracleDbType.Varchar2);
            //param.Value = Session["user"];
            //cmds.Parameters.Add(param);

            using (SqlDataReader ora = cmds.ExecuteReader())
            {

                while (ora.Read())
                {
                    //Label6.Text = odr["total"].ToString();
                    
                    Label20.Text = String.Format("{0:MM/dd/yyyy}", ora["FIR"]);
                }
            }
        }
    }

    //YenChang 06/30/2016 記錄每個item的Comment
    private void CommentRecord()
    {
        //GridView2.AllowPaging = false;
        Global.lb_rnum_gd2.Clear();
        Global.txComment_gd2.Clear();
        Label lb_rnum_gd2;
        TextBox tbComment_gd2;
        int test =GridView2.Rows.Count;
        int index =-1;
        List<String> gv2ronum_values =new List<string>();
        List<String> gv2comment =new List<string>();
        //bool[] values = (bool[])Session["page" + GridView1.PageIndex];
        if (Session["gv2ronum"] != null && Session["gv2comment"]!=null)
        {
             gv2ronum_values = (List<String>)Session["gv2ronum"];
             gv2comment = (List<String>)Session["gv2comment"];
             //index = gv2ronum_values.IndexOf(lb_rnum_gd2.Text);
        }
        else
        {
        }
        for (int i = 0; i < GridView2.Rows.Count;i++)
        {
            #region 記錄哪列改變
            tbComment_gd2 = (TextBox)GridView2.Rows[i].FindControl("TextBox44");
            GridViewRow row = (GridViewRow)tbComment_gd2.NamingContainer;
            lb_rnum_gd2 = row.Cells[5].FindControl("Label1") as Label;
            tbComment_gd2 = row.Cells[8].FindControl("TextBox44") as TextBox;
            if (gv2ronum_values != null)
            {
                index = gv2ronum_values.IndexOf(lb_rnum_gd2.Text);
            }
            #endregion
            //if (tbComment_gd2.ToString() != null)
            if (tbComment_gd2.Text != "")
            {
                //#region 記錄哪列改變
                //GridViewRow row = (GridViewRow)tbComment_gd2.NamingContainer;
                //#endregion
                Global.lb_rnum_gd2.Add(lb_rnum_gd2.Text);
                Global.txComment_gd2.Add(tbComment_gd2.Text);

            }
            else if (index != null && index != -1 )
            {
                Global.lb_rnum_gd2.Add(lb_rnum_gd2.Text);
                
                if (gv2comment!=null)
                {
                    Global.txComment_gd2.Add(gv2comment[index]);
                }

            }
            //GridView2.AllowPaging = true;
        }
        //GridView2.AllowPaging = true;
    }
   
    //參考網址:https://dotblogs.com.tw/mis2000lab/2012/08/07/rowdatabound_rowcreated_20120807
    protected void GridView2_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        //e.Row.Cells[5].Visible = true;
        Label lb_rnum_gd2;
        TextBox tbComment_gd2;
       if (e.Row.RowType == DataControlRowType.DataRow)
       {  //-- 當 GridView呈現「每一列」資料列（記錄）的時候，才會執行這裡！
            //-- 所以這裡就像迴圈一樣，會反覆執行喔！！

            //******************************************************
            if(Global.txComment_gd2.Count!=0)
            {
                //string rownum = e.Row.Cells[5].Text;
                lb_rnum_gd2 = e.Row.Cells[5].FindControl("Label1") as Label;
                int index = Global.lb_rnum_gd2.IndexOf(lb_rnum_gd2.Text);
                if (index != -1)
                {
                    //e.Row.Cells[6].Text = Global.txComment_gd2[index];
                    
                    tbComment_gd2 = e.Row.Cells[6].FindControl("TextBox44") as TextBox;
                    tbComment_gd2.Text = Global.txComment_gd2[index];
                }
                else
                {
                }
                //e.Row.Cells[4].ForeColor = System.Drawing.Color.Red;  
                //-- 把第五格的資料列（記錄）"格子"，變成紅色。
                //e.Row.Cells[4].Font.Bold = true;

            }

        }
        //e.Row.Cells[5].Visible = false;
    }

    //郵件標題,主收件人,副本,密件副本 by YenChang
    //參考網址https://social.msdn.microsoft.com/Forums/vstudio/en-US/8bc6e424-7dfc-43ab-b8a7-5c031d8ebcb9/c-how-to-return-a-liststring?forum=csharpgeneral
    private List<string> mail_get()
    {

        string sQUSTION_TYPE = "", sMAIL_TITLE = "", sMAIL_SEND_CC = "", sMAIL_SEND_GMTC = "", sMAIL_SEND_IT = "", sEMAIL = "", sEMAIL_CC = "";
        //取郵件TITLE與密件副本成員
        String sqlstr_mail_title = @"select QUSTION_TYPE,MAIL_TITLE,MAIL_SEND_CC,MAIL_SEND_GMTC,MAIL_SEND_IT  from dbo.OV_TRACE_QUSTION_DEF where UNI_NO='" + Session["user"].ToString() + "' and QUSTION_TYPE='" + DropDownList2.Text + "'";
        //正本收件人
        String sqlstr_mail_To = @"select EMAIL from dbo.pal_ec_apply where UNI_NO='" + Session["user"].ToString() + "'";
        //副本
        //String sqlstr_mail_CC = @"select EMAIL from dbo.pal_ec_apply where UNI_NO in ('"++"')";

        var retList = new List<string>();
        string connectionString = "Data Source=EC2;Initial Catalog=EC2;User ID=EC2;PASSWORD=gmtc";
        using (SqlConnection conns = new SqlConnection(connectionString))
        {
            conns.Open();


            //設定查詢
            SqlCommand cmds = conns.CreateCommand();
            //cmds.CommandText = "SELECT COUNT(*) AS DCOUNT FROM OV_CERT_DOWNLOAD WHERE CUS_ID = @CUS_ID " + Global.sqlstring2;

            cmds.CommandText = sqlstr_mail_title;

            using (SqlDataReader ora = cmds.ExecuteReader())
            {

                while (ora.Read())
                {
                    //預設只判斷一筆
                    if (sQUSTION_TYPE == "" || sMAIL_TITLE == "" || sMAIL_SEND_CC == "" || sMAIL_SEND_GMTC == "" || sMAIL_SEND_IT=="")
                    {
                        sQUSTION_TYPE += ora["QUSTION_TYPE"].ToString();
                        sMAIL_TITLE += ora["MAIL_TITLE"].ToString();
                        sMAIL_SEND_CC += ora["MAIL_SEND_CC"].ToString();
                        sMAIL_SEND_GMTC += ora["MAIL_SEND_GMTC"].ToString();
                        sMAIL_SEND_IT += ora["MAIL_SEND_IT"].ToString();
                    }
                    else
                    {
                        sQUSTION_TYPE = sQUSTION_TYPE + "," + ora["QUSTION_TYPE"].ToString();
                        sMAIL_TITLE = sMAIL_TITLE + "," + ora["MAIL_TITLE"].ToString();
                        sMAIL_SEND_CC = sMAIL_SEND_CC + "," + ora["MAIL_SEND_CC"].ToString();
                        sMAIL_SEND_GMTC = sMAIL_SEND_GMTC + "," + ora["MAIL_SEND_GMTC"].ToString();
                        sMAIL_SEND_IT = sMAIL_SEND_IT + ora["MAIL_SEND_IT"].ToString();

                    }
                    //retList[3] = ora["MAIL_SEND_CC"].ToString();
                    //sMail_content = sMail_content + "\n" + ora["TOSHIBA_PROJECT"].ToString().PadRight(25) + "\t" + ora["GMTC_PO"] + "\t" + ora["ITEM"] + "\t" + ora["WIP_LOT"] + "\t" + Global.txComment_gd2[index];
                }
                //宗富表示目前業務均為副本收件人3
                //sMAIL_SEND_CC = sMAIL_SEND_CC +"," + sMAIL_SEND_GMTC;
                retList.Add(sQUSTION_TYPE);//0
                retList.Add(sMAIL_TITLE);//1
                retList.Add(sMAIL_SEND_CC);//副本收件人姓名2
                //retList.Add(sMAIL_SEND_GMTC);//副本收件人 宗富表示目前業務均為副本收件人3
                retList.Add(sMAIL_SEND_IT);//3
            }


            cmds.CommandText = sqlstr_mail_To;
            using (SqlDataReader ora2 = cmds.ExecuteReader())
            {
                while (ora2.Read())
                {
                    if (sEMAIL == "")
                    {
                        sEMAIL += ora2["EMAIL"].ToString();

                    }
                    else
                    {
                        sEMAIL = sEMAIL + "," + ora2["EMAIL"].ToString();
                    }
                }

                
                if (sEMAIL == "")
                {
                    //增加宗富
                    retList.Add("liang@gmtc.com.tw,tom.tsai@gmtc.com.tw");//收件人
                }
                else
                {
                    retList.Add(sEMAIL);//收件人
                }
            }
            cmds.CommandText = sqlstr_mail_To;
            //副本收件人Email
            String sqlstr_mail_CC = @"select EMAIL from dbo.pal_ec_apply where UNI_NO in (" + retList[2] + ")";
            using (SqlDataReader ora3 = cmds.ExecuteReader())
            {
                while (ora3.Read())
                {
                    if (sEMAIL_CC == "")
                    {
                        sEMAIL_CC += ora3["EMAIL"].ToString();

                    }
                    else
                    {
                        sEMAIL_CC = sEMAIL_CC + "," + ora3["EMAIL"].ToString();
                    }
                }


                if (sEMAIL_CC == "")
                {
                    //增加宗富
                    retList.Add("liang@gmtc.com.tw,tom.tsai@gmtc.com.tw");//副本收件人
                }
                else
                {
                    sEMAIL_CC = sEMAIL_CC + "," + sMAIL_SEND_GMTC;
                    retList.Add(sEMAIL_CC);//副本收件人5
                }
            }
            }
        //GridView2.Visible = true;
        return retList;
    }



    protected void CheckBox4_CheckedChanged(object sender, EventArgs e)
    {
        CheckedAllBox();
        if (CheckBox4.Checked == true)
        {
            CheckBox5.Checked = false;
        }
        CheckboxallRecord();
        ContentQry();
    }

    protected void CheckBox5_CheckedChanged1(object sender, EventArgs e)
    {
        CheckedAllBox();
        if (CheckBox5.Checked == true)
        {
            CheckBox4.Checked = false;
        }
        CheckboxallRecord();
        ContentQry();
    }
    //WIP顯示(milestone), 當URGENT_STATUS不為空白時, 本資料GMTC_PO字體秀紅色
    public List<string> GMTC_Red()
    {
        int index=0;
        //本資料GMTC_PO字體秀紅色
        List<String> gmtc_po = new List<String>();
        string connectionString = "Data Source=EC2;Initial Catalog=EC2;User ID=EC2;PASSWORD=gmtc";
        using (SqlConnection conns = new SqlConnection(connectionString))
        {
            conns.Open();
            string sql = @"select distinct TSB_PO  TSB_PO from dbo.OV_ORDER_TRACE_JAP100 where row_num>0 and URGENT_STATUS!=''
                           and GMTC_PO  in (select GMTC_PO   from dbo.OV_ORDER_TRACE_WIP )";
            //設定查詢
            SqlCommand cmds = conns.CreateCommand();
            //cmds.CommandText = "SELECT COUNT(*) AS DCOUNT FROM OV_CERT_DOWNLOAD WHERE CUS_ID = @CUS_ID " + Global.sqlstring2;

            cmds.CommandText = sql;

            using (SqlDataReader ora = cmds.ExecuteReader())
            {

                while (ora.Read())
                {
                    gmtc_po.Add(ora["TSB_PO"].ToString());
                        index++;
                }
            }
        }
        return gmtc_po;
    }
    protected void GridView1_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        LinkButton lb_tsb_po;
        Label lb_tsb_po_2;
        LinkButton lb_item;
        Label lb_item2;
        Label rownum;
        List<String> tsb_po = new List<String>();
        List<String> item = new List<String>();
        tsb_po = GMTC_Red();
        item = ITEM_Red();
        if (e.Row.RowType == DataControlRowType.DataRow)
        {  //-- 當 GridView呈現「每一列」資料列（記錄）的時候，才會執行這裡！
            //-- 所以這裡就像迴圈一樣，會反覆執行喔！！

            //******************************************************
            lb_tsb_po = e.Row.Cells[7].FindControl("LinkButton6") as LinkButton;
            lb_tsb_po_2 = e.Row.Cells[7].FindControl("Label5") as Label;
            lb_item = e.Row.Cells[3].FindControl("LinkButton5") as LinkButton;
            lb_item2 = e.Row.Cells[3].FindControl("Label66") as Label;
            rownum = e.Row.Cells[24].FindControl("Label7") as Label;
            if (tsb_po.Contains(lb_tsb_po.Text))
            {
                //string rownum = e.Row.Cells[5].Text;
                //lb_gmtc_po = e.Row.Cells[2].FindControl("Label6") as Label;
                //整列變色
                //e.Row.ForeColor = System.Drawing.Color.Red; //文字顏色改為紅色
                //某欄變色
                //e.Row.Cells[2].ForeColor = System.Drawing.Color.Red;
                lb_tsb_po.BackColor = System.Drawing.Color.Red;
                lb_tsb_po_2.Visible = false;
            }
            else
            {
                if (lb_tsb_po_2 != null && lb_tsb_po != null)
                {
                    lb_tsb_po_2.Visible = true;
                    lb_tsb_po.Visible = false;
                }
            }
            if (item.Contains(rownum.Text) && Session["user"].ToString() != "JAP10021")
            {
                lb_item.BackColor = System.Drawing.Color.Red;
                lb_item2.Visible = false;
            }
            else
            {
                lb_item.Visible = false;
                lb_item2.Visible = true;

            }
        }
    }
    //改由TSB_PO呈現
    protected void LinkButton4_Click(object sender, EventArgs e)
    {
        #region 記錄哪列改變
        LinkButton pb = (LinkButton)sender;
        GridViewRow row = (GridViewRow)pb.NamingContainer;
        #endregion
        LinkButton pblabel = (row.Cells[2].FindControl("LinkButton4") as LinkButton);
        Session["gmtc_po"] = pblabel.Text;
        this.Response.Write("<script language=javascript>window.open('Detail_Track.aspx','Detail','toolbar=yes,scrollbars=yes,resizable=yes,top=100,left=200,width=650,height=600')</script>");

    }
    
    protected void LinkButton5_Click(object sender, EventArgs e)
    {
       // Response.Redirect("IFRAME.htm");
        //Button31.Click(this, e);
        #region 記錄哪列改變
        LinkButton pb = (LinkButton)sender;
        GridViewRow row = (GridViewRow)pb.NamingContainer;
        #endregion
        LinkButton pblabel = (row.Cells[2].FindControl("LinkButton4") as LinkButton);
        Label pblabel2 = (row.Cells[5].FindControl("Label66") as Label);
        Session["gmtc_po"] = pblabel.Text;
        Session["item"] = pblabel2.Text;
        this.Response.Write("<script language=javascript>window.open('Note.aspx','Note','toolbar=yes,scrollbars=yes,resizable=yes,top=100,left=200,width=750,height=600')</script>");

        //Button31_Click(this, e);
    }
    
    protected void Button31_Click(object sender, EventArgs e)
    {

    }
    protected void LinkButton6_Click(object sender, EventArgs e)
    {
        #region 記錄哪列改變
        LinkButton pb = (LinkButton)sender;
        GridViewRow row = (GridViewRow)pb.NamingContainer;
        #endregion
        LinkButton pblabel = (row.Cells[2].FindControl("LinkButton4") as LinkButton);
        LinkButton pbitem = (row.Cells[3].FindControl("LinkButton5") as LinkButton);
        LinkButton pblabel2 = (row.Cells[7].FindControl("LinkButton6") as LinkButton);
        Session["gmtc_po"] = pblabel.Text;
        Session["tsb_po"] = pblabel2.Text;
        Session["item"] = pbitem.Text;
        this.Response.Write("<script language=javascript>window.open('Detail_Track.aspx','Detail','toolbar=yes,scrollbars=yes,resizable=yes,top=100,left=200,width=650,height=600')</script>");
    }

    //判斷FIC_COMMENT是否有值
    public List<string> ITEM_Red()
    {
        int index = 0;
        //本資料GMTC_PO字體秀紅色
        List<String> item = new List<String>();
        string connectionString = "Data Source=EC2;Initial Catalog=EC2;User ID=EC2;PASSWORD=gmtc";
        using (SqlConnection conns = new SqlConnection(connectionString))
        {
            conns.Open();
            string sql = @"select distinct row_num rownum from  dbo.OV_ORDER_TRACE_JAP100 where note != '' or fic_comment != ''";
            //設定查詢
            SqlCommand cmds = conns.CreateCommand();
            //cmds.CommandText = "SELECT COUNT(*) AS DCOUNT FROM OV_CERT_DOWNLOAD WHERE CUS_ID = @CUS_ID " + Global.sqlstring2;

            cmds.CommandText = sql;

            using (SqlDataReader ora = cmds.ExecuteReader())
            {

                while (ora.Read())
                {
                    item.Add(ora["rownum"].ToString());
                    index++;
                }
            }
        }
        return item;
    }
}