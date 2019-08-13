using System;
using System.Linq;
using System.Web.Http;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using Microsoft.ProjectOxford.Face;
using Microsoft.ProjectOxford.Face.Contract;

namespace photoRollCall_1.Controllers
{
    public class LineBotController : ApiController
    {
        string sqlcn = "server=TINA\\SQLEXPRESS;database=rollCallSystem;UID=SQLAdmin;PWD=1234";

        bool IDExist;
        int memType;
        int adminID;
        int studentID;
        string personID;
        string prevDT = "";
        string nextDT = "";
        int prevD = -1;
        int nowD = -1;
        int nextD = -1;
        int picID;

        [HttpPost]
        public async System.Threading.Tasks.Task<IHttpActionResult> POSTAsync()
        {
            FaceServiceClient faceSC = new FaceServiceClient("9ed1e677b9bd402c83697c7c9e31e3e7",
                                    "https://eastasia.api.cognitive.microsoft.com/face/v1.0");
            string ChannelAccessToken = Properties.Settings.Default.ChannelAccessToken;
            string groupID = "group1";
            string replyToken = "";

            //回覆訊息
            string Message = "無法辨識的指令，請使用功能選項輸入關鍵字!";

            try
            {
                //剖析JSON
                string postData = Request.Content.ReadAsStringAsync().Result;
                var ReceivedMessage = isRock.LineBot.Utility.Parsing(postData);
                replyToken = ReceivedMessage.events[0].replyToken;

                using (SqlConnection cnn1 = new SqlConnection(sqlcn))
                {
                    using (SqlCommand cmd1 = new SqlCommand())
                    {
                        cmd1.Connection = cnn1;
                        cmd1.CommandText = "SELECT A.memType, A.adminID, A.studentID, B.personID FROM dbo.members AS A, dbo.students AS B WHERE A.userID = @userId AND A.studentID = B.studentID";
                        cmd1.Parameters.Add("@userId", SqlDbType.NVarChar).Value = ReceivedMessage.events[0].source.userId;

                        cnn1.Open();
                        SqlDataReader rr = cmd1.ExecuteReader();
                        IDExist = rr.HasRows;
                        if (IDExist)
                        {
                            rr.Read();
                            memType = int.Parse(rr[0].ToString());
                            adminID = int.Parse(rr[1].ToString());
                            studentID = int.Parse(rr[2].ToString());
                            personID = rr[3].ToString();
                        }

                        rr.Close();
                        cmd1.Dispose();
                        cnn1.Close();
                    }
                }

                using (SqlConnection cnn11 = new SqlConnection(sqlcn))
                {
                    using (SqlCommand cmd11 = new SqlCommand())
                    {
                        cmd11.Connection = cnn11;
                        cmd11.CommandText = "SELECT * FROM dbo.classes";

                        cnn11.Open();
                        SqlDataReader rr = cmd11.ExecuteReader();
                        if (rr.HasRows)
                        {
                            while (rr.Read())
                            {
                                TimeSpan timeSpan = DateTime.Now.Subtract(DateTime.Parse(rr[1].ToString()));
                                if (DateTime.Parse(rr[1].ToString()).ToShortDateString() == DateTime.Now.ToShortDateString())
                                    nowD = int.Parse(rr[0].ToString());
                                else if (timeSpan.TotalDays > 0)
                                {
                                    prevD = int.Parse(rr[0].ToString());
                                    prevDT = DateTime.Parse(rr[1].ToString()).ToShortDateString();
                                }
                                else
                                {
                                    nextD = int.Parse(rr[0].ToString());
                                    nextDT = DateTime.Parse(rr[1].ToString()).ToShortDateString();
                                    break;
                                }
                            }
                        }

                        rr.Close();
                        cmd11.Dispose();
                        cnn11.Close();
                    }
                }

                //判斷初次加入好友
                if (ReceivedMessage.events.FirstOrDefault().type == "follow")
                {
                    if (!IDExist)
                    {
                        using (SqlConnection cnn2 = new SqlConnection(sqlcn))
                        {
                            using (SqlCommand cmd2 = new SqlCommand())
                            {
                                cmd2.Connection = cnn2;
                                cmd2.CommandText = "INSERT INTO dbo.members(userID, memType) VALUES(@userId, 0)";
                                cmd2.Parameters.Add("@userId", SqlDbType.NVarChar).Value = ReceivedMessage.events[0].source.userId;
                                cnn2.Open();

                                if (cmd2.ExecuteNonQuery() == 1)
                                    Message = "登錄LineBot帳號成功!\n請輸入您的ID(學生輸入學號 如：1091234，老師輸入註冊代碼)：";
                                else
                                {
                                    Message = "登錄LineBot帳號失敗!\n請刪除此LineBot後重新加入或找教授/助教確認個人資料。";
                                }
                                cmd2.Dispose();
                                cnn2.Close();
                            }
                        }
                    }
                    else
                    {
                        Message = "您的userID已被註冊過!\n請隨意傳送資料或找教授/助教確認個人資料。";
                    }

                    isRock.LineBot.Utility.ReplyMessage(replyToken, Message, ChannelAccessToken);
                    return Ok();
                }
                else if (ReceivedMessage.events.FirstOrDefault().type == "message")
                {
                    if (ReceivedMessage.events.FirstOrDefault().message.type.Trim().ToLower() == "text")
                    {
                        if (memType == 0)
                        {
                            if (ReceivedMessage.events[0].message.text.Length != 7)
                                Message = "輸入的ID字數錯誤，ID為7個字的字串!";
                            else
                            {
                                using (SqlConnection cnn3 = new SqlConnection(sqlcn))
                                {
                                    using (SqlCommand cmd3 = new SqlCommand())
                                    {
                                        cmd3.Connection = cnn3;
                                        cmd3.CommandText = "SELECT studentID FROM dbo.students WHERE studentAct = @studentAct";
                                        cmd3.Parameters.Add("@studentAct", SqlDbType.NVarChar).Value = ReceivedMessage.events[0].message.text;

                                        cnn3.Open();
                                        SqlDataReader rr = cmd3.ExecuteReader();
                                        if (rr.HasRows)
                                        {
                                            memType = 2;
                                            rr.Read();
                                            studentID = int.Parse(rr[0].ToString());
                                        }

                                        rr.Close();
                                        cmd3.Dispose();
                                        cnn3.Close();
                                    }
                                }

                                if (memType == 0)
                                {
                                    using (SqlConnection cnn3 = new SqlConnection(sqlcn))
                                    {
                                        using (SqlCommand cmd3 = new SqlCommand())
                                        {
                                            cmd3.Connection = cnn3;
                                            cmd3.CommandText = "SELECT adminID FROM dbo.admins WHERE adminAct = @adminAct";
                                            cmd3.Parameters.Add("@adminAct", SqlDbType.NVarChar).Value = ReceivedMessage.events[0].message.text;

                                            cnn3.Open();
                                            SqlDataReader rr = cmd3.ExecuteReader();
                                            if (rr.HasRows)
                                            {
                                                memType = 1;
                                                rr.Read();
                                                adminID = int.Parse(rr[0].ToString());
                                            }

                                            rr.Close();
                                            cmd3.Dispose();
                                            cnn3.Close();
                                        }
                                    }
                                }

                                if (memType > 0)
                                {

                                    if (memType == 1)
                                    {
                                        using (SqlConnection cnn4 = new SqlConnection(sqlcn))
                                        {
                                            using (SqlCommand cmd4 = new SqlCommand())
                                            {
                                                cmd4.Connection = cnn4;
                                                CreatePersonResult pr = await faceSC.CreatePersonAsync(groupID, studentID.ToString());
                                                cmd4.CommandText = "UPDATE dbo.admins SET adminExist = 1 WHERE adminID = @adminID";
                                                cmd4.Parameters.Add("@adminID", SqlDbType.Int).Value = adminID;
                                                cnn4.Open();

                                                cmd4.ExecuteNonQuery();

                                                cmd4.Dispose();
                                                cnn4.Close();
                                            }
                                        }
                                    }
                                    else if (memType == 2)
                                    {
                                        using (SqlConnection cnn4 = new SqlConnection(sqlcn))
                                        {
                                            using (SqlCommand cmd4 = new SqlCommand())
                                            {
                                                cmd4.Connection = cnn4;
                                                CreatePersonResult pr = await faceSC.CreatePersonAsync(groupID, studentID.ToString());
                                                cmd4.CommandText = "UPDATE dbo.students SET personID = @personID, studentExist = 1 WHERE studentID = @studentID";
                                                cmd4.Parameters.Add("@personID", SqlDbType.NVarChar).Value = pr.PersonId.ToString();
                                                cmd4.Parameters.Add("@studentID", SqlDbType.Int).Value = studentID;
                                                cnn4.Open();

                                                cmd4.ExecuteNonQuery();

                                                cmd4.Dispose();
                                                cnn4.Close();
                                            }
                                        }
                                    }

                                    using (SqlConnection cnn4 = new SqlConnection(sqlcn))
                                    {
                                        using (SqlCommand cmd4 = new SqlCommand())
                                        {
                                            cmd4.Connection = cnn4;
                                            cmd4.CommandText = "UPDATE dbo.members SET adminID = @adminID, studentID = @studentID, memType = @memType WHERE userID = @userId";
                                            cmd4.Parameters.Add("@adminID", SqlDbType.Int).Value = adminID;
                                            cmd4.Parameters.Add("@studentID", SqlDbType.Int).Value = studentID;
                                            cmd4.Parameters.Add("@memType", SqlDbType.Int).Value = memType;
                                            cmd4.Parameters.Add("@userId", SqlDbType.NVarChar).Value = ReceivedMessage.events[0].source.userId;
                                            cnn4.Open();

                                            if (cmd4.ExecuteNonQuery() == 1)
                                            {
                                                if (memType == 1)
                                                    Message = "個人資訊註冊完畢\n感謝您的填寫\n\n您可以開始傳送照片點名及使用功能選項中的查詢點名結果確認出席狀況";
                                                else if (memType == 2)
                                                    Message = "請傳送臉部可清晰辨識五官的個人照(還需三張)：";
                                            }
                                            else
                                            {
                                                Message = "登入ID失敗!\n請重新輸入或找教授/助教確認個人資料。";
                                            }
                                            cmd4.Dispose();
                                            cnn4.Close();
                                        }
                                    }
                                }
                            }
                        }
                        else if(ReceivedMessage.events[0].message.text == "[線上請假]")
                        {
                            if(nextD == -1)
                            {
                                Message = "本課程已經結束!";
                                isRock.LineBot.Utility.ReplyMessage(replyToken, Message, ChannelAccessToken);
                                return Ok();
                            }
                            if (memType == 5)
                            {
                                using (SqlConnection cnn10 = new SqlConnection(sqlcn))
                                {
                                    using (SqlCommand cmd10 = new SqlCommand())
                                    {
                                        cmd10.Connection = cnn10;
                                        cmd10.CommandText = "INSERT INTO dbo. rollCall(studentID, classID, RCState, picID, RCTime) VALUES(@studentID, @classID, 2, 1, @RCTime)";
                                        cmd10.Parameters.Add("@studentID", SqlDbType.Int).Value = studentID;
                                        cmd10.Parameters.Add("@classID", SqlDbType.Int).Value = nextD;
                                        cmd10.Parameters.Add("@RCTime", SqlDbType.DateTime).Value = DateTime.Now;
                                        cnn10.Open();

                                        if (cmd10.ExecuteNonQuery() != 1)
                                            Message = "請假失敗!\n請重新傳送或找教授/助教確認個人資料。";
                                        else
                                            Message = "下一堂課時間為 " + nextDT + "\n\n已請假成功!";
                                        cmd10.Dispose();
                                        cnn10.Close();
                                    }
                                }
                            }
                            else if(memType == 1)
                            {
                                using (SqlConnection cnn8 = new SqlConnection(sqlcn))
                                {
                                    using (SqlCommand cmd8 = new SqlCommand())
                                    {
                                        cmd8.Connection = cnn8;
                                        cmd8.CommandText = "SELECT DISTINCT A.studentAct FROM dbo.students AS A, dbo.rollCall AS B WHERE A.studentID = B.studentID AND B.RCState = 2 AND B.classID = @classID";
                                        cmd8.Parameters.Add("@classID", SqlDbType.Int).Value = nextD;
                                        cnn8.Open();

                                        SqlDataReader rr = cmd8.ExecuteReader();
                                        if (!rr.HasRows)
                                            Message = "下一堂課時間為 " + nextDT + "\n\n今天無人請假!";
                                        else
                                        {
                                            Message = "下一堂課時間為 " + nextDT + "\n\n請假學生：";
                                            while (rr.Read())
                                                Message +=  "\n" + rr[0].ToString();
                                        }
                                        rr.Close();
                                        cmd8.Dispose();
                                        cnn8.Close();
                                    }
                                }
                            }
                        }
                        else if (ReceivedMessage.events[0].message.text == "[查詢點名結果]")
                        {
                            if (prevD == -1 && nowD == -1)
                            {
                                Message = "本課程還沒開始!";
                                isRock.LineBot.Utility.ReplyMessage(replyToken, Message, ChannelAccessToken);
                                return Ok();
                            }
                            if (memType == 5)
                            {
                                string temp = "";
                                using (SqlConnection cnn12 = new SqlConnection(sqlcn))
                                {
                                    using (SqlCommand cmd12 = new SqlCommand())
                                    {
                                        cmd12.Connection = cnn12;
                                        cmd12.CommandText = "SELECT RCState, RCTime FROM dbo.rollCall WHERE studentID = @studentID AND classID = @classID ORDER BY RCID DESC";
                                        cmd12.Parameters.Add("@studentID", SqlDbType.Int).Value = studentID;
                                        if (nowD == -1)
                                        {
                                            cmd12.Parameters.Add("@classID", SqlDbType.Int).Value = prevD;
                                            Message = "最近一次點名紀錄為\n" + prevDT;
                                        }
                                        else
                                        {
                                            cmd12.Parameters.Add("@classID", SqlDbType.Int).Value = nowD;
                                            Message = "最近一次點名紀錄為\n" + DateTime.Now.ToShortDateString();
                                        }
                                        cnn12.Open();
                                        SqlDataReader rr = cmd12.ExecuteReader();
                                        if (!rr.HasRows)
                                            temp = "缺席";
                                        else
                                        {
                                            rr.Read();
                                            Message += " " + DateTime.Parse(rr[1].ToString()).ToShortTimeString();
                                            if (rr[0].ToString() == "2")
                                                temp = "請假";
                                            else if (rr[0].ToString() == "1")
                                                temp = "出席";
                                            else
                                                temp = "缺席";
                                        }
                                        Message += "\n\n您的點名紀錄為： " + temp;

                                        rr.Close();
                                        cmd12.Dispose();
                                        cnn12.Close();
                                    }
                                }
                            }
                            else if (memType == 1)
                            {
                                using (SqlConnection cnn8 = new SqlConnection(sqlcn))
                                {
                                    using (SqlCommand cmd8 = new SqlCommand())
                                    {cmd8.Connection = cnn8;
                                        cmd8.CommandText = "SELECT DISTINCT A.studentAct, B.RCState FROM (SELECT *  FROM dbo.students WHERE studentID > 1) AS A LEFT OUTER JOIN dbo.rollCall AS B ON A.studentID = B.studentID AND B.classID = @classID ORDER BY B.RCState DESC";
                                        if (nowD == -1)
                                        {
                                            cmd8.Parameters.Add("@classID", SqlDbType.Int).Value = prevD;
                                            Message = "最近一次點名紀錄為\n" + prevDT;
                                        }
                                        else
                                        {
                                            cmd8.Parameters.Add("@classID", SqlDbType.Int).Value = nowD;
                                            Message = "最近一次點名紀錄為\n" + DateTime.Now.ToShortDateString();
                                        }
                                        cnn8.Open();

                                        SqlDataReader rr = cmd8.ExecuteReader();
                                        bool check = true;

                                        Message += "\n\n請假：";
                                        check = rr.Read();
                                        while (check && rr[1].ToString() == "2")
                                        {
                                            Message += "\n" + rr[0].ToString();
                                            check = rr.Read();
                                        }
                                        Message += "\n\n出席：";
                                        while (check && rr[1].ToString() == "1")
                                        {
                                            Message += "\n" + rr[0].ToString();
                                            check = rr.Read();
                                        }
                                        Message += "\n\n缺席：";
                                        while (check)
                                        {
                                            Message += "\n" + rr[0].ToString();
                                            check = rr.Read();
                                        }

                                        rr.Close();
                                        cmd8.Dispose();
                                        cnn8.Close();
                                    }
                                }
                            }
                        }
                        else if(ReceivedMessage.events[0].message.text == "[使用說明]")
                        {
                            Message = "●查詢點名結果：\n" + 
                                      "學生可以看到自己最近一次上課出席狀況\n" +
                                      "老師可以看到最近一堂課學生出缺勤狀況\n\n" +
                                      "●線上請假：\n" +
                                      "學生可以為下一次的課程請假\n" +
                                      "老師可以看到下一次上課請假學生資訊\n\n" +
                                      "●管理網頁：\n" +
                                      "此功能學生不適用\n" +
                                      "老師點選可直接連結到點名後台，手動更改點名資訊\n\n" +
                                      "●智慧點名方式：\n" +
                                      "此功能學生不適用\n" +
                                      "老師拍攝並傳送教室學生照片，系統會透過人臉辨識完成點名";
                        }
                        /*else if (ReceivedMessage.events[0].message.text == "[管理網站]")
                        {
                            Message = "http://140.138.155.175:8080/index.aspx";
                        }*/

                        isRock.LineBot.Utility.ReplyMessage(replyToken, Message, ChannelAccessToken);
                        return Ok();
                    }
                    else if (ReceivedMessage.events.FirstOrDefault().message.type.Trim().ToLower() == "image")
                    {
                        if (memType >= 1 &&memType <= 4)
                        {
                            //取得contentid
                            var LineContentID = ReceivedMessage.events.FirstOrDefault().message.id.ToString();
                            //取得bytedata
                            var filebody = isRock.LineBot.Utility.GetUserUploadedContent(LineContentID, ChannelAccessToken);

                            var ms = new MemoryStream(filebody);
                            Face[] faces = await faceSC.DetectAsync(ms);

                            string picPath = "/Temp/" + Guid.NewGuid() + ".jpg";
                            var path = System.Web.HttpContext.Current.Request.MapPath(picPath);
                            //上傳圖片
                            File.WriteAllBytes(path, filebody);

                            using (SqlConnection cnn15 = new SqlConnection(sqlcn))
                            {
                                using (SqlCommand cmd15 = new SqlCommand())
                                {
                                    cmd15.Connection = cnn15;
                                    cmd15.CommandText = "INSERT INTO dbo.pictures(studentID, picPath, picTime) VALUES(@studentID, @picPath, @picTime)";
                                    cmd15.Parameters.Add("@studentID", SqlDbType.Int).Value = studentID;
                                    cmd15.Parameters.Add("@picPath", SqlDbType.VarChar).Value = picPath;
                                    cmd15.Parameters.Add("@picTime", SqlDbType.DateTime).Value = DateTime.Now;
                                    cnn15.Open();

                                    cmd15.ExecuteNonQuery();

                                    cmd15.Dispose();
                                    cnn15.Close();
                                }
                            }

                            if (memType != 1)
                            {
                                if (faces.Length != 1)
                                {
                                    Message = "無法辨識，請重新傳送五官清晰的個人照!";
                                    isRock.LineBot.Utility.ReplyMessage(replyToken, Message, ChannelAccessToken);
                                    return Ok();
                                }

                                //上傳圖片
                                memType ++;
                                using (SqlConnection cnn7 = new SqlConnection(sqlcn))
                                {
                                    using (SqlCommand cmd7 = new SqlCommand())
                                    {
                                        cmd7.Connection = cnn7;
                                        cmd7.CommandText = "UPDATE dbo.members SET memType = @memType WHERE studentID = @studentID";
                                        cmd7.Parameters.Add("@memType", SqlDbType.Int).Value = memType;
                                        cmd7.Parameters.Add("@studentID", SqlDbType.Int).Value = studentID;
                                        cnn7.Open();

                                        if (cmd7.ExecuteNonQuery() != 1)
                                            Message = "登入圖片失敗!\n請重新傳送或找教授/助教確認個人資料。";
                                        cmd7.Dispose();
                                        cnn7.Close();
                                    }
                                }
                                MemoryStream stream = new MemoryStream(filebody);
                                AddPersistedFaceResult result = await faceSC.AddPersonFaceAsync(groupID, Guid.Parse(personID), stream);
                                await faceSC.TrainPersonGroupAsync(groupID);

                                if (memType == 3)
                                    Message = "請傳送臉部可清晰辨識五官的個人照(還需兩張)：";
                                else if (memType == 4)
                                    Message = "請傳送臉部可清晰辨識五官的個人照(還需一張)：";
                                else if(memType == 5)
                                    Message = "個人資訊註冊完畢\n感謝您的填寫\n\n您可以開始使用功能選項中的查詢點名功能及線上請假功能";
                            }
                            else
                            {
                                if(nowD == -1)
                                {
                                    Message = "今日沒有上課喔!";
                                    isRock.LineBot.Utility.ReplyMessage(replyToken, Message, ChannelAccessToken);
                                    return Ok();
                                }

                                using (SqlConnection cnn16 = new SqlConnection(sqlcn))
                                {
                                    using (SqlCommand cmd16 = new SqlCommand())
                                    {
                                        cmd16.Connection = cnn16;
                                        cmd16.CommandText = "SELECT picID FROM dbo.pictures WHERE picPath = @picPath";
                                        cmd16.Parameters.Add("@picPath", SqlDbType.NVarChar).Value = picPath;

                                        cnn16.Open();
                                        SqlDataReader rr = cmd16.ExecuteReader();
                                        if (rr.HasRows)
                                        {
                                            rr.Read();
                                            picID = int.Parse(rr[0].ToString());
                                        }

                                        cmd16.Dispose();
                                        cnn16.Close();
                                    }
                                }

                                // 將照片中的臉，與指定的PersonGroup進行比對
                                if (faces != null)
                                {
                                    Message = "已偵測到 " + faces.Length + " 人\n";
                                    // 將這張照片回傳的人臉，取出每一張臉的FaceId並進行轉換成Guid
                                    Guid[] faceGuids = faces.Select(x => x.FaceId).ToArray();
                                    int pCount = 0;

                                    if (faceGuids.Length > 0)
                                    {
                                        // 透過Identify，找出在這張照片中，所有辨識出的人臉，是否有包含在PersonGroup中的所有人員
                                        IdentifyResult[] result = await faceSC.IdentifyAsync(groupID, faceGuids);
                                        // 取得照片中在PersonGroup裡的人
                                        for (int i = 0; i < result.Length; i++)
                                        {
                                            for (int p = 0; p < result[i].Candidates.Length; p++)
                                            {
                                                Person person = await faceSC.GetPersonAsync(groupID, result[i].Candidates[p].PersonId);
                                                string strPersonId = person.Name;
                                                pCount++;

                                                using (SqlConnection cnn9 = new SqlConnection(sqlcn))
                                                {
                                                    using (SqlCommand cmd9 = new SqlCommand())
                                                    {
                                                        cmd9.Connection = cnn9;
                                                        cmd9.CommandText = "INSERT INTO dbo. rollCall(studentID, classID, RCState, picID, RCTime) VALUES(@studentID, @classID, 1, @picID, @RCTime)";
                                                        cmd9.Parameters.Add("@studentID", SqlDbType.Int).Value = int.Parse(strPersonId);
                                                        cmd9.Parameters.Add("@classID", SqlDbType.Int).Value = nowD;
                                                        cmd9.Parameters.Add("@picID", SqlDbType.Int).Value = picID;
                                                        cmd9.Parameters.Add("@RCTime", SqlDbType.DateTime).Value = System.DateTime.Now;
                                                        cnn9.Open();

                                                        if (cmd9.ExecuteNonQuery() != 1)
                                                            Message = "點名失敗，請再嘗試一次!";
                                                        cmd9.Dispose();
                                                        cnn9.Close();
                                                    }
                                                }
                                            }
                                        }
                                        Message += "成功辨識 " + pCount + " 人";
                                    }
                                }
                            }

                            //ReplyMessage
                            isRock.LineBot.Utility.ReplyMessage(replyToken, Message, ChannelAccessToken);
                            return Ok();
                        }
                        else
                        {
                            isRock.LineBot.Utility.ReplyMessage(replyToken, "else", ChannelAccessToken);
                            return Ok();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                isRock.LineBot.Utility.ReplyMessage(replyToken, "錯誤訊息： " + ex.Message + "\n請聯絡工程人員!", ChannelAccessToken);
                return Ok();
            }
            return Ok();
        }
    }
}
