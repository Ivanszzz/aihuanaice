using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Security;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace UploadNaiceDevice
{
    /// <summary>
    /// 设备状态
    /// </summary>
    public class DevState
    {
        public string DeviceNo;
        public string DeviceTime;
        public string DeviceState;
    }

    /// <summary>
    /// 服务器返回消息
    /// </summary>
    public class ClientResponseMsg
    {
        public bool IsSuccess { get; set; }
        public string Messaage { get; set; }
        public object Data { get; set; }
        public int totalRows { get; set; }
    }

    /// <summary>
    /// 奈测机台状态上传
    /// </summary>
    public class NaiceDeviceState
    {
        public String ConnectionString { get; set; }

        public String Start()
        {
            return UploadDeviceState();
        }

        /// <summary>
        /// 为INI文件中指定的节点取得字符串
        /// </summary>
        /// <param name="lpAppName">欲在其中查找关键字的节点名称</param>
        /// <param name="lpKeyName">欲获取的项名</param>
        /// <param name="lpDefault">指定的项没有找到时返回的默认值</param>
        /// <param name="lpReturnedString">指定一个字串缓冲区，长度至少为nSize</param>
        /// <param name="nSize">指定装载到lpReturnedString缓冲区的最大字符数量</param>
        /// <param name="lpFileName">INI文件完整路径</param>
        /// <returns>复制到lpReturnedString缓冲区的字节数量，其中不包括那些NULL中止字符</returns>
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, int nSize, string lpFileName);

        /// <summary>
        /// 读取INI文件值
        /// </summary>
        /// <param name="section">节点名</param>
        /// <param name="key">键</param>
        /// <param name="def">未取到值时返回的默认值</param>
        /// <param name="filePath">INI文件完整路径</param>
        /// <returns>读取的值</returns>
        public static string Read(string section, string key, string def, string filePath)
        {
            StringBuilder sb = new StringBuilder(1024);
            GetPrivateProfileString(section, key, def, sb, 1024, filePath);
            return sb.ToString();
        }

        /// <summary>
        /// 检查文件是否存在
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        private static void CheckPath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                throw new ArgumentNullException("filePath");
            }
        }

        /// <summary>
        /// 获取XML指定节点下的值
        /// </summary>
        /// <param name="sourceName"></param>
        /// <returns></returns>
        public string GetXML(string sourceName)
        {
            
            string xmlFileName = System.AppDomain.CurrentDomain.BaseDirectory +"/UploadConfig.xml";
            string result = "";
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFileName);
            //获取到xml文件的根节点
            XmlElement nodeRoot = xmlDoc.DocumentElement;

            //UploadConfig.xml，查看指定节点的值
            XmlNodeList uploadConfig = nodeRoot.SelectNodes("*");
            foreach (XmlNode config in uploadConfig)
            {
                if (config.FirstChild.InnerText.Equals("001"))
                {
                    XmlNodeList configChidNodeList = config.SelectNodes("*");
                    foreach (XmlNode configChileNode in configChidNodeList)
                    {
                        if (configChileNode.Name.Equals(sourceName))
                        {
                            result = configChileNode.InnerText + "";
                        }
                    }
                    Console.WriteLine();
                    break;
                }
                else
                {
                    continue;
                }
            }
            return result;
        }

        /// <summary>
        /// 客户端请求数据
        /// </summary>
        public static string ClientRequest(string actionUrl, string jsonBody)
        {
            string returnMsg = string.Empty;
            ClientResponseMsg respMsg = new ClientResponseMsg();
            returnMsg = JsonConvert.SerializeObject(respMsg);
            if (!string.IsNullOrEmpty(jsonBody))
            {
                returnMsg = HttpPost(actionUrl, jsonBody.ToString(), "application/json", "POST", "");
            }
            return returnMsg;
        }

        #region POST
        public static string HttpPost(string url, string body, string contentType = "application/json", string Method = "POST", string bearerToken = "")
        {

            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(url);

            httpWebRequest.ContentType = contentType;
            httpWebRequest.Method = Method;
            if (!string.IsNullOrEmpty(bearerToken))
            {
                httpWebRequest.Headers.Add(HttpRequestHeader.Authorization, "Bearer " + bearerToken);
            }
            httpWebRequest.Timeout = 200000;

            if (Method == "POST" || Method == "PUT")
            {
                byte[] btBodys = Encoding.UTF8.GetBytes(body);
                httpWebRequest.ContentLength = btBodys.Length;
                httpWebRequest.GetRequestStream().Write(btBodys, 0, btBodys.Length);
            }

            string responseContent = string.Empty;

            try
            {
                using (HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse())
                {
                    using (Stream zipstream = httpWebResponse.GetResponseStream())
                    {
                        if (zipstream != null)
                        {
                            using (StreamReader reader = new StreamReader(zipstream, System.Text.Encoding.GetEncoding("utf-8")))
                            {
                                responseContent = reader.ReadToEnd();
                            }
                        }
                    }

                }
                httpWebRequest.Abort();
            }
            catch (WebException ex)
            {
                var httpRspn = (HttpWebResponse)ex.Response;
                throw ex;
            }
            return responseContent;
        }
        #endregion

        /// <summary>
        /// 上传机台状态数据
        /// </summary>
        public string UploadDeviceState()
        {
            DevState devState = new DevState();
            string filePath = GetXML("IniPath");
            CheckPath(filePath);
            string devicemsg = Read("DevState", "DevState", null, filePath);
            if (devicemsg == null)
            {
                return "file is empty";
            }

            devState.DeviceNo = devicemsg.Split(',')[0];
            devState.DeviceTime = devicemsg.Split(',')[1];
            devState.DeviceState = devicemsg.Split(',')[2];

            string url = GetXML("MESURL");
            string returnMsg = string.Empty;
            JObject reqData = new JObject();
            reqData.Add("DeviceNo", devState.DeviceNo);
            reqData.Add("DeviceTime", devState.DeviceTime);
            reqData.Add("DeviceState", devState.DeviceState);
            string body = JsonConvert.SerializeObject(reqData);
            returnMsg = ClientRequest(url, body);

            ClientResponseMsg msgObj = JsonConvert.DeserializeObject<ClientResponseMsg>(returnMsg);
            if (!msgObj.IsSuccess)
            {
                return devState.DeviceNo + "DeviceState Upload fail";
            }
            else
            {
                return devState.DeviceNo + "DeviceState Upload success";
            }
        }

    }

    /// <summary>
    /// 奈测CSV数据上传
    /// </summary>
    public class NaiceDeviceCSV
    {

        public String ConnectionString { get; set; }

        public String Start()
        {
            return UploadCSV();
        }

        /// <summary>
        /// 上传文件
        /// </summary>
        /// <param name="url">请求地址</param>
        /// <param name="fs">文件流</param>
        /// <param name="fileName">文件名</param>
        /// <returns></returns>
        public static string HttpUploadFile(string url, String path, string fileName, String[] keys, String[] values)
        {
            using (var client = new HttpClient())
            {
                using (var multipartFormDataContent = new MultipartFormDataContent())
                {
                    int len = keys.Length;
                    for (int i = 0; i < len; i++)
                    {
                        client.DefaultRequestHeaders.Add(keys[i], values[i]); // 存在HEAD里面
                    }
                    multipartFormDataContent.Add(new StreamContent(File.Open(path, FileMode.Open)), "file", fileName);
                    Task<HttpResponseMessage> message = client.PostAsync(url, multipartFormDataContent);
                    message.Wait();
                    MemoryStream ms = new MemoryStream();
                    Task t = message.Result.Content.CopyToAsync(ms);
                    t.Wait();
                    ms.Position = 0;
                    StreamReader sr = new StreamReader(ms, Encoding.UTF8);
                    string content = sr.ReadToEnd();
                    return content;
                }
            }
        }

        /// <summary>
        /// 上传文件成功后修改文件后缀名
        /// </summary>
        /// <param name="filepath">文件夹</param>
        /// <param name="filename">文件名</param>
        /// <returns></returns>
        public void changeFileName(string filepath, string filename)
        {
            File.Move(filepath + filename, filepath + filename + ".bak");
        }

        /// <summary>
        /// 复制文件到文件夹
        /// </summary>
        /// <param name="path">最终文件路径</param>
        /// <param name="fileName">指定文件的完整路径</param>
        public void saveFile(string path, string fileName)
        {
            if (File.Exists(fileName))//必须判断要复制的文件是否存在
            {
                File.Copy(fileName, path, true);//三个参数分别是源文件路径，存储路径，若存储路径有相同文件是否替换
            }
        }

        /// <summary>
        /// 将文件夹下的所有文件copy到当前文件夹下的temp目录
        /// </summary>
        /// <param name="filepath">文件夹路径</param>
        public void copyFileToTemp(string filepath)
        {
            //创建文件夹tmp
            DirectoryInfo di = Directory.CreateDirectory(filepath + "//temp");
            DirectoryInfo folder = new DirectoryInfo(filepath);

            foreach (FileInfo file in folder.GetFiles("*.csv"))
            {
                string name = file.FullName;//带路径的名称
                string filename = Path.GetFileName(file.ToString());
                saveFile(filepath + "temp/" + filename, name);
            }
        }

        /// <summary>
        /// 获取XML指定节点下的值
        /// </summary>
        /// <param name="sourceName"></param>
        /// <returns></returns>
        public string GetXML(string sourceName)
        {
            string xmlFileName = "UploadConfig.xml";
            string result = "";
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(xmlFileName);
            //获取到xml文件的根节点
            XmlElement nodeRoot = xmlDoc.DocumentElement;

            //UploadConfig.xml，查看指定节点的值
            XmlNodeList uploadConfig = nodeRoot.SelectNodes("*");
            foreach (XmlNode config in uploadConfig)
            {
                if (config.FirstChild.InnerText.Equals("001"))
                {
                    XmlNodeList configChidNodeList = config.SelectNodes("*");
                    foreach (XmlNode configChileNode in configChidNodeList)
                    {
                        if (configChileNode.Name.Equals(sourceName))
                        {
                            result = configChileNode.InnerText + "";
                        }
                    }
                    Console.WriteLine();
                    break;
                }
                else
                {
                    continue;
                }
            }
            return result;
        }

        /// <summary>
        /// 上传CSV文件到服务器
        /// </summary>
        public string UploadCSV()
        {

            string url = GetXML("URL");
            string path = GetXML("CSVPath");
            //将目录下的所有文件复制到temp目录
            copyFileToTemp(path);

            DirectoryInfo folder = new DirectoryInfo(path);

            //找到CSV目录下最近修改的文件文件名
            Dictionary<string, DateTime> dict = new Dictionary<string, DateTime>();
            foreach (FileInfo file in folder.GetFiles("*.csv"))
            {
                if((Path.GetFileNameWithoutExtension(file.ToString()).Split('-').Length - 1) != 3)
                {
                    continue;
                }
                string filename = Path.GetFileName(file.ToString());//仅文件名
                DateTime updateTime = file.LastWriteTime;
                dict.Add(filename, updateTime);
            }
            if(dict.Keys.Count == 0)
            {
                return "can not find file in '" + path + "'";
            }
            string maxkey = dict.Keys.Select(x => new { x, y = dict[x] }).OrderByDescending(x => x.y).First().x;

            //上传除最近修改的文件的所有文件
            foreach (FileInfo file in folder.GetFiles("*.csv"))
            {
                if ((Path.GetFileNameWithoutExtension(file.ToString()).Split('-').Length - 1) != 3)
                {
                    continue;
                }
                string filepath = file.FullName;
                string filename = Path.GetFileName(file.ToString());
                if (maxkey == filename)
                {
                    continue;
                }
                if (HttpUploadFile(url, filepath, filename, new String[] { "source", "sortCol" ,"tableName","savePath"}, new String[] { GetXML("DeviceNo"), null, GetXML("TableName"), GetXML("savePath") }) == "Success")
                {
                    File.Delete(filepath);//上传成功后删除文件
                }
                else
                {
                    continue;
                    //return "Upload "+ filepath + " fail";
                }
            }
            string filepath_ = GetXML("CSVPath") + "temp/" + maxkey;
            //上传 temp目录下,最近修改的文件
            if(HttpUploadFile(url, filepath_, maxkey, new String[] { "source", "sortCol", "tableName", "savePath" }, new String[] { GetXML("DeviceNo"), null, GetXML("TableName"), GetXML("savePath") })!= "Success")
            {
                return "Upload '" + filepath_ + "' fail";
            }
            return "Upload '"+ filepath_ + "' success";
        }
    }

}
