using System;
using System.IO;
using System.Net;
using System.Runtime.Serialization.Json;
using System.Xml;
using System.Text;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace WecomMessageIntegration
{
    class WecomHelper
    {
        #region 发送消息需要的信息
        // 企业ID
        private string m_corpId = string.Empty;
        // 应用的凭证密钥
        private string m_secret = string.Empty;
        // 应用id
        private string m_agentId = string.Empty;
        // 调用企业微信API接口的凭证
        private string m_accessToken = string.Empty;
        // 上次获取accessToken的时间
        private string m_tokenTime = string.Empty;
        #endregion
        /// <summary>
        /// 获得xml文件的路径
        /// </summary>
        /// <returns>xml文件路径</returns>
        private string GetXmlPath()
        {
            return Path.Combine(Environment.CurrentDirectory, "WecomSetting.xml");
        }
        /// <summary>
        /// 从xml配置文件中初始化信息
        /// </summary>
        private void InitConfig()
        {
            var strXmlPathFile = GetXmlPath();
            if (!File.Exists(strXmlPathFile))
            {
                throw new Exception($"【WecomSetting】配置不存在:{strXmlPathFile}");
            }
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(strXmlPathFile);    //加载Xml文件  
            XmlNode root = xmlDoc.SelectSingleNode("wecom");
            foreach (XmlElement node in root)
            {
                switch (node.Name)
                {
                    case "corpid":
                        m_corpId = node.GetAttribute("value");
                        break;
                    case "secret":
                        m_secret = node.GetAttribute("value");
                        break;
                    case "agentid":
                        m_agentId = node.GetAttribute("value");
                        break;
                    case "accesstoken":
                        m_accessToken = node.GetAttribute("value");
                        m_tokenTime = node.GetAttribute("lasttime");
                        break;
                }

            }
        }
        /// <summary>
        /// 初始化accessToken
        /// </summary>
        private void InitAccessToken()
        {
            #region 检测目前缓存的token是否过期，官方规定是7200s(2小时)，未防止误差，设置7000s
            if (m_accessToken != null && m_accessToken != string.Empty && m_tokenTime != null && m_tokenTime != string.Empty)
            {
                var lastTokenTime = Convert.ToDateTime(m_tokenTime);
                var time = DateTime.Now - lastTokenTime;
                if (time.TotalSeconds <= 7000) return;
            }
            #endregion
            #region 获取accessToken
            var tokenUrl = $"https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={m_corpId}&corpsecret={m_secret}";
            WebResponse result = null;
            WebRequest req = WebRequest.Create(tokenUrl);
            result = req.GetResponse();
            Stream s = result.GetResponseStream();
            XmlDictionaryReader xmlReader = JsonReaderWriterFactory.CreateJsonReader(s, XmlDictionaryReaderQuotas.Max);
            xmlReader.Read();
            string xml = xmlReader.ReadOuterXml();
            s.Close();
            s.Dispose();
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(xml);
            XmlElement rootElement = doc.DocumentElement;
            m_accessToken = rootElement.SelectSingleNode("access_token").InnerText.Trim();
            #endregion
            #region 写入新的accessToken和tokenTime
            string strXmlPathFile = GetXmlPath();
            if (!File.Exists(strXmlPathFile))
            {
                throw new Exception(string.Format("【WecomSetting】配置不存在:{0}", strXmlPathFile));
            }

            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(strXmlPathFile);    //加载Xml文件  
            XmlNode root = xmlDoc.SelectSingleNode("wecom");
            foreach (XmlElement node in root)
            {
                if (node.Name == "accesstoken")
                {
                    node.SetAttribute("value", m_accessToken);
                    node.SetAttribute("lasttime", DateTime.Now.ToString());
                }
            }
            xmlDoc.Save(strXmlPathFile);
            #endregion
        }
        /// <summary>
        /// 通过手机号获得userid
        /// </summary>
        /// <param name="strMobiles">手机号，以"|"分隔</param>
        /// <returns>userid，以"|"分隔</returns>
        private IEnumerable<string> GetUserIdByMobile(string strMobiles)
        {
            var url = $"https://qyapi.weixin.qq.com/cgi-bin/user/getuserid?access_token={m_accessToken}";
            var lstMobile = strMobiles.Split("|");
            foreach (var mobile in lstMobile)
            {
                var req = WebRequest.Create(url);
                req.ContentType = "application/x-www-form-urlencoded";
                req.Method = "POST";

                var param = new
                {
                    mobile
                };

                Byte[] b = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(param));
                req.ContentLength = b.Length;
                Stream stream = req.GetRequestStream();
                stream.Write(b, 0, b.Length);
                WebResponse result = req.GetResponse();
                Stream s = result.GetResponseStream();
                XmlDictionaryReader xmlReader = JsonReaderWriterFactory.CreateJsonReader(s, XmlDictionaryReaderQuotas.Max);
                xmlReader.Read();
                string xml = xmlReader.ReadOuterXml();
                s.Close();
                s.Dispose();
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(xml);
                XmlElement rootElement = doc.DocumentElement;
                yield return rootElement.SelectSingleNode("userid").InnerText.Trim();
            }
        }
        /// <summary>
        /// 获取接受消息的user
        /// </summary>
        /// <returns></returns>
        private string GetUserList()
        {
            return string.Join("|", GetUserIdByMobile("16657119235"));
        }
        /// <summary>
        /// 获取发送的消息内容
        /// </summary>
        /// <returns>消息内容</returns>
        private string GetMessageContent()
        {
            var strb = new StringBuilder();
            strb.Append("测试消息（请勿回复）：\n");
            strb.Append("不知你看到这条消息，明白消息发送的过程了吗？");
            return strb.ToString();
        }
        /// <summary>
        /// 发送消息
        /// </summary>
        /// <returns>发送结果</returns>
        public string SendMessage()
        {
            InitConfig();
            InitAccessToken();
            string retString = string.Empty;
            string url = $"https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token={m_accessToken}";//string url = GetUrlString;
            HttpWebRequest request = WebRequest.Create(url) as HttpWebRequest;
            request.Method = "POST";
            request.ContentType = "application/json";
            var sendContent = new
            {
                touser = GetUserList(),
                msgtype = "text",
                agentid = m_agentId,
                text = new
                {
                    content = GetMessageContent()
                }
            };
            var strContent = JsonConvert.SerializeObject(sendContent);
            using (StreamWriter dataStream = new StreamWriter(request.GetRequestStream()))
            {
                dataStream.Write(strContent);
                dataStream.Close();
            }
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            string encoding = response.ContentEncoding;
            if (encoding == null || encoding.Length < 1)
            {
                encoding = "UTF-8"; //默认编码  
            }
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.GetEncoding(encoding));
            retString = reader.ReadToEnd();
            return retString;
        }

    }
}
