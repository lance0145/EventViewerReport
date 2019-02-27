using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Xml;
using System.Text.RegularExpressions;
using System.Security.Cryptography;

namespace EventViewerReport
{
    class Utility
{
    public DataTable DatabaseConfig = null;
    private static string defaultUserName = "WPLGlobalUser";
    private static string defaultPassword = "2RleqdHjMLseKHFzRNq4Bg****";

    private string smtpAccount = "";
    public string SmtpAccount
    {
        get { return smtpAccount; }
        set { smtpAccount = value; }
    }
    private string smtpPassword = "";
    public string SmtpPassword
    {
        get { return smtpPassword; }
        set { smtpPassword = value; }
    }
    private string smtpServer = "";
    public string SmtpServer
    {
        get { return smtpServer; }
        set { smtpServer = value; }
    }
    private string smptPort = "";
    public string SmptPort
    {
        get { return smptPort; }
        set { smptPort = value; }
    }
    private string logFileBaseLocation = "";
    public string LogFileBaseLocation
    {
        get { return logFileBaseLocation; }
        set { logFileBaseLocation = value; }
    }

    public void writeToOAILogFile(string input)
    {
        string currentDate = DateTime.Now.ToString("yyyyMMdd");
        string fileName = "logs/" + currentDate + "_OAILogFile.txt";

        writeToFileWithTime(fileName, input);
    }
    public void writeToXMLLogFile(string input)
    {
        string currentDate = DateTime.Now.ToString("yyyyMMdd");
        string fileName = "logs/" + currentDate + "_XMLLogFile.txt";

        writeToFileWithTime(fileName, input);
    }
    public void writeToHTMLPullLogFile(string input)
    {
        string currentDate = DateTime.Now.ToString("yyyyMMdd");
        string fileName = "logs/" + currentDate + "_HTMLPullLogFile.txt";

        writeToFileWithTime(fileName, input);
    }
    public void writeToXMLPullLogFile(string input)
    {
        string currentDate = DateTime.Now.ToString("yyyyMMdd");
        string fileName = "logs/" + currentDate + "_XMLPulLogFile.txt";

        writeToFileWithTime(fileName, input);
    }
    public void writeToHTMLstoparse(string input)
    {
        string currentDate = DateTime.Now.ToString("yyyyMMdd");
        string fileName = "logs/" + currentDate + "_HTMLstoparseLogFile.txt";

        writeToFileWithTime(fileName, input);
    }
    //==========================================================================================

    public void writeToFileWithTime(string fileName, string input)
    {
        WriteToFile(fileName, input, true);
    }
    public void writeToFile(string fileName, string input)
    {
        WriteToFile(fileName, input, false);
    }
    public void WriteToFile(string fileName, string input, bool isIncludeTime)
    {
        string currentFileName = fileName;
        if (string.IsNullOrEmpty(currentFileName))
        {
            currentFileName = getCurrentLogFileName();
        }

        try
        {
            string currentFilePath = Path.GetDirectoryName(currentFileName);
            if (!string.IsNullOrEmpty(currentFilePath))
            {
                Directory.CreateDirectory(currentFilePath);
            }
            else
            {
                currentFileName = @".\" + currentFileName;
            }
            using (StreamWriter w = File.AppendText(fileName))
            {
                if (isIncludeTime)
                {
                    w.WriteLine("{0}: {1}"
                            , DateTime.Now.ToString()
                            , input
                        );
                }
                else
                {
                    w.WriteLine("{0}"
                            , input
                        );
                }
                w.Flush();
                w.Close();
            }
        }
        catch { }
    }

    public string getCurrentLogFileName(bool isIncludeDirectory = true, string additionalLabel = "")
    {
        string returnValue = "";

        string currentDate = DateTime.Now.ToString("yyyyMMdd");
        if (isIncludeDirectory)
        {
            returnValue = "logs/log_" + currentDate;
        }
        else
        {
            returnValue = "log_" + currentDate;
        }
        if (!string.IsNullOrEmpty(additionalLabel))
        {
            returnValue += "_" + additionalLabel;
        }

        returnValue += ".txt";
        return returnValue;
    }
    public void logAction(string activity)
    {
        logAction(activity, true, "");
    }
    public void logAction(string activity, string additionalLabel)
    {
        logAction(activity, true, additionalLabel);
    }
    public void logAction(string activity, bool isLabelUpdated, string additionalLabel)
    {
        string pathToLogFile = getCurrentLogFileName(true, additionalLabel);
        if (!string.IsNullOrEmpty(logFileBaseLocation))
        {
            //Z:\Utilities\Logs\<MachineName>\log_20171231_29.txt
            pathToLogFile = logFileBaseLocation;
            if (!pathToLogFile.EndsWith("\""))
            {
                pathToLogFile += @"\";
            }
            pathToLogFile += getCurrentLogFileName(false, additionalLabel);
        }

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(pathToLogFile));
            using (StreamWriter w = File.AppendText(pathToLogFile))
            {
                w.WriteLine("{0}: {1}"
                        , DateTime.Now.ToString()
                        , activity
                    );
                w.Flush();
                w.Close();
            }

            if (
                activity.ToUpper().Contains("ERROR:")
                || activity.ToUpper().Contains("EXCEPTION:")
            )
            {
                string pathToErrorFile = getCurrentLogFileName(true, "").Replace("log_", "ERROR_" + System.Reflection.Assembly.GetExecutingAssembly().GetName().Name);
                using (StreamWriter w = File.AppendText(pathToErrorFile))
                {
                    w.WriteLine("{0}: {1}"
                            , DateTime.Now.ToString()
                            , activity
                        );
                    w.Flush();
                    w.Close();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("EXCEPTION: Could not write to file : " + activity + " : " + ex.ToString());
        }
    }

    public string GetXmlDbPathWplData(string type, string pathToRoot, string value)
    {
        string returnValue = "";

        try
        {
            string path = pathToRoot + @"\WplData\" + type;
            if (GetDoesPathExist(path, true))
            {
                returnValue = path + @"\" + value + ".xml";
            }
        }
        catch { }

        return returnValue;
    }

    public string GetXmlDbPathCms(string type, string pathToRoot, string value)
    {
        string returnValue = "";

        try
        {
            string path = pathToRoot + @"\Cms\" + type;
            if (GetDoesPathExist(path, true))
            {
                returnValue = path + @"\" + value + ".xml";
            }
        }
        catch { }

        return returnValue;
    }

    public string GetXmlDbPathSite(string type, string pathToRoot, string value)
    {
        string returnValue = "";

        try
        {
            string path = pathToRoot + @"\Site\" + type;
            if (GetDoesPathExist(path, true))
            {
                returnValue = path + @"\" + value + ".xml";
            }
        }
        catch { }

        return returnValue;
    }

    public string GetXmlDbPathCmsHeader(string type, string pathToRoot, string siteType, string siteName, string isUserAdmin, string isUserAffiliateAdmin, string isUserContentEditor, string isSgAdmin, string isUserRepAgencyAdmin)
    {
        string returnValue = "";

        try
        {
            string path = pathToRoot + @"\Cms\" + type;

            if (GetDoesPathExist(path, true))
            {
                returnValue = path + @"\" + siteType + "_" + siteName + "_" + isUserAdmin + "_" + isUserAffiliateAdmin + "_" + isUserContentEditor + "_" + isSgAdmin + "_" + isUserRepAgencyAdmin + ".xml";
            }
        }
        catch { }

        return returnValue;
    }

    public string GetXmlDbPathCmsHeader(string type, string pathToRoot, string inclusionKey, string isUserAdmin, string isUserAffiliateAdmin, string isUserContentEditor, string isSgAdmin, string isUserRepAgencyAdmin)
    {
        string returnValue = "";

        try
        {
            string path = pathToRoot + @"\Cms\" + type;

            if (GetDoesPathExist(path, true))
            {
                returnValue = path + @"\" + inclusionKey + "_" + isUserAdmin + "_" + isUserAffiliateAdmin + "_" + isUserContentEditor + "_" + isSgAdmin + "_" + isUserRepAgencyAdmin + ".xml";
            }
        }
        catch { }

        return returnValue;
    }

    public string GetXmlDbPath(string type, string pathToRoot, string id)
    {
        return GetXmlDbPath(type, pathToRoot, id, id);
    }

    public string GetXmlDbPath(string type, string pathToRoot, string id, string fileName)
    {
        string returnValue = "";
        int folderSize = 200000;
        double multiplier = 0;
        string folderValue = "";

        try
        {
            double dblId = Convert.ToInt64(id);
            multiplier = Math.Ceiling(dblId / folderSize);

            folderValue = Convert.ToString(multiplier * folderSize);
            string path = pathToRoot + @"\WplData\" + type + @"\" + folderValue;
            if (GetDoesPathExist(path, true))
            {
                returnValue = path + @"\" + fileName + ".xml";
            }
        }
        catch { }

        return returnValue;
    }

    public string GetXmlDbPathWheArticleTitle(string pathToRoot, string part1, string part2, string fileName, string language)
    {
        string returnValue = "";

        try
        {
            string path = pathToRoot + @"\Whe\ArticleTitle\" + language + @"\" + part1 + @"\" + part2;
            if (GetDoesPathExist(path, true))
            {
                returnValue = path + @"\" + fileName + ".xml";
            }
        }
        catch { }

        return returnValue;
    }

    public bool GetDoesPathExist(string path, bool createIfMissing)
    {
        bool retunValue = System.IO.Directory.Exists(path);

        if (!retunValue && createIfMissing)
        {
            System.IO.Directory.CreateDirectory(path);
            retunValue = true;
        }
        return retunValue;
    }

    public string getXmlEncodedString(string input)
    {
        string decode = WebUtility.HtmlDecode(input);
        string encode = WebUtility.HtmlEncode(decode);

        StringBuilder builder = new StringBuilder();
        foreach (char c in encode)
        {
            if ((int)c > 127)
            {
                builder.Append("&#");
                builder.Append((int)c);
                builder.Append(";");
            }
            else if ((int)c <= 0)
            {
            }
            else
            {
                builder.Append(c);
            }
        }
        string returnResult = builder.ToString();

        return returnResult;
    }

    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    //////// New Db initalization calls
    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    // Load all Database connections from local XML file, if any data is missing it can crash the whole process
    public void LoadDatabases(string xmlFileName = "DbConfig.xml")
    {
        if (DatabaseConfig == null)
        {
            LoadDatabases(xmlFileName, true);
        }
    }
    public void LoadDatabases(string xmlFileName, bool isForcedReload)
    {
        if (DatabaseConfig == null)
        {
            isForcedReload = true;
        }

        if (isForcedReload)
        {
            try
            {
                XmlTextReader reader = new XmlTextReader(xmlFileName);

                DataSet dsDatabaseConfig = new DataSet();
                dsDatabaseConfig.ReadXml(reader);

                DatabaseConfig = dsDatabaseConfig.Tables[0];
            }
            catch (Exception e)
            {
                logAction("LoadDatabases: EXCEPTION: Could not loadDatabaseConnections - " + e.ToString());
            }
        }
    }
    // Iterate through all databases from the xml until the server and database combo is found, then send the connection string to the other method
    public SqlConnection InitializeConnectionForDatabase(string databaseName, string ipAddress)
    {
        return InitializeConnectionForDatabase(databaseName, ipAddress, defaultUserName, Decrypt(defaultPassword));
    }
    public SqlConnection InitializeConnectionForDatabase(string databaseName, string ipAddress, string userName, string password)
    {
        SqlConnection connection = null;

        string connectionString = GetDatabaseConnectionString(databaseName, ipAddress, userName, password);

        if (!string.IsNullOrEmpty(connectionString))
        {
            try
            {
                connection = InitializeDatabaseConnection(connectionString);

                if (connection != null)
                {
                    return connection;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception e)
            {
                logAction("InitializeConnectionForDatabase: EXCEPTION: Could not initialize DB " + databaseName + " at " + ipAddress + " - " + e.ToString());
            }
        }
        else
        {
            logAction("InitializeConnectionForDatabase: ERROR: Could not initialize DB " + databaseName + " at " + ipAddress + " - Connection string was empty");
        }

        return null;
    }
    public string GetDatabaseConnectionString(string databaseName, string ipAddress, string userName, string password)
    {
        string connectionString = "";

        try
        {
            if (DatabaseConfig != null)
            {
                for (int i = 0; i < DatabaseConfig.Rows.Count; i++)
                {
                    DataRow dr = DatabaseConfig.Rows[i];

                    if (dr["Database"].ToString().ToLower().Trim() == databaseName.ToLower().Trim())
                    {
                        if (Convert.ToBoolean(Convert.ToInt32(dr["IsConnectionGood"].ToString())))
                        {
                            connectionString = dr["ConnectionString"].ToString().Trim().Replace("REPLACEIPADDRESS", ipAddress).Replace("REPLACEUSERNAME", userName).Replace("REPLACEPASSWORD", password);
                        }

                        if (string.IsNullOrEmpty(connectionString))
                        {
                            dr["IsConnectionGood"] = "0";
                        }
                    }
                }
            }
        }
        catch (Exception e)
        {
            logAction("GetDatabaseConnectionString: ERROR: Could not create connection string - " + e.ToString());
        }

        return connectionString;
    }

    // Create connection object and open it based on the connection string, logs exceptions
    public SqlConnection InitializeDatabaseConnection(string connectionString)
    {
        try
        {
            if (!string.IsNullOrEmpty(connectionString))
            {
                SqlConnection connection = new SqlConnection(connectionString);
                return connection;
            }
        }
        catch (Exception e)
        {
            logAction("InitializeDatabaseConnection: ERROR: Could not initialize DB " + connectionString + " - " + e.ToString());
        }

        return null;
    }

    public bool IsTimeOfDayBetween(DateTime time, TimeSpan startTime, TimeSpan endTime)
    {
        if (endTime == startTime)
        {
            return true;
        }
        else if (endTime < startTime)
        {
            return time.TimeOfDay <= endTime || time.TimeOfDay >= startTime;
        }
        else
        {
            return time.TimeOfDay >= startTime && time.TimeOfDay <= endTime;
        }

    }

    public string GetPropertyValue(string key)
    {
        string propertyFileName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name + ".properties";
        return GetPropertyValue(key, propertyFileName, "");
    }
    public string GetPropertyValue(string key, string propertyFileName, string defaultReturnValue = "")
    {
        string returnValue = "";
        FileInfo propertyFile = new FileInfo(propertyFileName);
        if (File.Exists(propertyFile.FullName))
        {
            //"WplServerCheck.properties"
            System.IO.StreamReader file = new System.IO.StreamReader(propertyFile.FullName);
            string line;
            while ((line = file.ReadLine()) != null)
            {
                string[] parts = line.Split('=');
                if (parts.Length > 1)
                {
                    try
                    {
                        if (parts[0].Trim().ToLower() == key.Trim().ToLower())
                        {
                            for (int i = 1; i < parts.Length; i++)
                            {
                                if (!string.IsNullOrEmpty(returnValue))
                                {
                                    returnValue += "=";
                                }
                                returnValue += parts[i].Trim();
                            }
                        }
                    }
                    catch { }
                }
            }
        }

        if (string.IsNullOrEmpty(returnValue))
        {
            returnValue = defaultReturnValue;
        }

        return returnValue;
    }

    public byte[] GetDecompressedGzipBytes(byte[] input)
    {
        using (var compressedStream = new MemoryStream(input))
        using (var zipStream = new GZipStream(compressedStream, CompressionMode.Decompress))
        using (var resultStream = new MemoryStream())
        {
            zipStream.CopyTo(resultStream);
            return resultStream.ToArray();
        }
    }

    public byte[] ConvertStringToBytes(string input)
    {
        byte[] bytes = new byte[input.Length * sizeof(char)];
        System.Buffer.BlockCopy(input.ToCharArray(), 0, bytes, 0, bytes.Length);
        return bytes;
    }

    public string ConvertBytesToString(byte[] input)
    {
        char[] chars = new char[input.Length / sizeof(char)];
        System.Buffer.BlockCopy(input, 0, chars, 0, input.Length);
        return new string(chars);
    }

    public Stream GenerateStreamFromString(string s)
    {
        MemoryStream stream = new MemoryStream();
        StreamWriter writer = new StreamWriter(stream);
        writer.Write(s);
        writer.Flush();
        stream.Position = 0;
        return stream;
    }

    public string GetSourceCodeFromUrl(string url)
    {
        string returnValue = "";
        UserProfile currentUserProfile = new UserProfile();
        logAction("GetSourceCodeFromUrl: INFO: Processing url " + url);

        try
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.UserAgent = currentUserProfile.UserAgent;
            request.Timeout = 800000; // 800 seconds
            request.ReadWriteTimeout = 800000; // 800 seconds
            request.ProtocolVersion = System.Net.HttpVersion.Version10;
            request.ServicePoint.ConnectionLimit = 10000;
            request.ServicePoint.Expect100Continue = false;
            request.Method = currentUserProfile.Method;
            request.CookieContainer = currentUserProfile.CookieJar;
            HttpWebResponse response;

            try
            {
                response = HttpWebResponseExt.GetResponseNoException(request);
                logAction("GetSourceCodeFromUrl: INFO: Data from GetResponse");
            }
            catch (Exception ex)
            {
                response = null;
                logAction("GetSourceCodeFromUrl: EXCEPTION: GetResponse error is " + ex.ToString());
            }

            // Get HTML from response if 
            if (string.IsNullOrEmpty(returnValue) && response != null)
            {
                if (
                    response.StatusCode == HttpStatusCode.OK ||
                    response.StatusCode == HttpStatusCode.Moved ||
                    response.StatusCode == HttpStatusCode.Redirect)
                {

                    // if the remote file was found, download oit
                    using (var inputStream = new StreamReader(response.GetResponseStream()))
                    {
                        returnValue = inputStream.ReadToEnd();
                    }
                    /*
                    using (Stream outputStream = File.OpenWrite(fileName))
                    {
                        byte[] buffer = new byte[4096];
                        int bytesRead;
                        do
                        {
                            bytesRead = inputStream.Read(buffer, 0, buffer.Length);
                            outputStream.Write(buffer, 0, bytesRead);
                        } while (bytesRead != 0);
                    }
                    */
                }
            }
            else if (string.IsNullOrEmpty(returnValue) && response == null)
            {
                logAction("GetSourceCodeFromUrl: EXCEPTION: from " + url + " : Exception : Response is null");
            }
        }
        catch (Exception ex)
        {
            logAction("GetSourceCodeFromUrl: EXCEPTION: from " + url + " : Exception : " + ex.ToString());
        }

        return returnValue;
    }

    public void printLog(string message)
    {
        DateTime now = DateTime.Now;
        Console.WriteLine(now + " : " + message);
    }

    public void saveFileTo(string filePath, string content, bool endWithNewLine = true)
    {
        try
        {
            FileInfo fileInfo = new FileInfo(filePath);
            bool dirExists = Directory.Exists(fileInfo.Directory.ToString());

            if (!dirExists)
            {
                Directory.CreateDirectory(fileInfo.Directory.ToString());
            }

            StreamWriter outFile = new StreamWriter(filePath);
            if (endWithNewLine)
            {
                outFile.WriteLine(contentToUTF8(content));
            }
            else
            {
                outFile.Write(contentToUTF8(content));
            }
            outFile.Close();
        }
        catch (Exception ex)
        {
            Console.WriteLine("EXEPTION: " + ex.ToString());
        }
    }

    public void appendToFile(string filePath, string content)
    {
        try
        {
            FileInfo fileInfo = new FileInfo(filePath);
            bool dirExists = Directory.Exists(fileInfo.Directory.ToString());

            if (!dirExists)
            {
                Directory.CreateDirectory(fileInfo.Directory.ToString());
            }

            if (!File.Exists(filePath))
            {
                saveFileTo(filePath, content, true);
            }
            else
            {
                StreamWriter sw = File.AppendText(filePath);
                sw.WriteLine(content);
                sw.Close();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("EXEPTION: " + ex.ToString());
        }
    }

    public string PathAddBackslash(string path)
    {
        Console.WriteLine("path = {0}", path);
        string separator1 = Path.DirectorySeparatorChar.ToString();
        string separator2 = Path.AltDirectorySeparatorChar.ToString();

        path = path.Trim();

        if (path.EndsWith(separator1) || path.EndsWith(separator2))
            return path;

        if (path.Contains(separator2))
            return path + separator2;

        Console.WriteLine("Path = {0}", path);

        return path + separator1;
    }

    public string contentToUTF8(string content)
    {
        string result = "";

        try
        {
            UTF8Encoding utf8 = new UTF8Encoding();
            Byte[] encodeBytes = utf8.GetBytes(content);
            result = utf8.GetString(encodeBytes);
        }
        catch (Exception ex)
        {
            Console.WriteLine("EXEPTION: " + ex.ToString());
        }

        return result;
    }

    public string getCurrentTimestampString()
    {
        return DateTime.Now.ToString("yyyy.MM.dd.HH.mm.ss.fff");
    }

    public void pressEnterToExit()
    {
        Console.Write("Press enter to exit...");
        Console.ReadLine();
        Environment.Exit(0);
    }

    public List<string> getFiles(string directoryPath, bool recursive = true)
    {
        List<string> result = new List<string>();

        try
        {
            string[] fileEntries = Directory.GetFiles(directoryPath);

            foreach (string fileName in fileEntries)
            {
                result.Add(fileName);
            }

            if (recursive)
            {
                string[] subdirectoryEntries = Directory.GetDirectories(directoryPath);
                foreach (string subdirectory in subdirectoryEntries)
                {
                    result = result.Concat(getFiles(subdirectory, recursive)).ToList();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("EXEPTION: " + ex.ToString());
        }

        return result;
    }

    public List<string> GetDirectories(string directoryPath, bool isRecursive = true, bool isTryAgain = true)
    {
        List<string> result = new List<string>();

        try
        {
            result.Add(directoryPath);

            if (isRecursive)
            {
                string[] subdirectoryEntries = Directory.GetDirectories(directoryPath);
                foreach (string subdirectory in subdirectoryEntries)
                {
                    result = result.Concat(GetDirectories(subdirectory, isRecursive)).ToList();
                }
            }
        }
        catch (Exception ex)
        {
            if (isTryAgain)
            {
                if (result.Contains(directoryPath))
                {
                    result.RemoveAt(result.Count - 1);
                }
                GetDirectories(directoryPath, isRecursive, false);
            }
            else
            {
                Console.WriteLine("EXCEPTION: Could not fully process at " + directoryPath + " : " + ex.ToString());
            }
        }

        return result;
    }

    public string ReadFile(string filePath)
    {
        string returnValue = "";

        try
        {
            returnValue = File.ReadAllText(filePath, Encoding.UTF8);
        }
        catch (Exception ex)
        {
            logAction("ReadFile: EXCEPTION: Could not read file at " + filePath + ": " + ex.ToString());
        }

        return returnValue;
    }

    public bool GetIsStringFoundInString(string value, string matchPhrase)
    {
        bool returnValue = false;

        Regex myRegex = new Regex(matchPhrase, RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.IgnorePatternWhitespace);
        Match match = myRegex.Match(value);
        if (match.Success)
        {
            returnValue = true;
        }

        return returnValue;
    }

    public string DecodeXml(string input)
    {
        string returnValue = "";

        if (!string.IsNullOrEmpty(input))
        {
            returnValue = System.Net.WebUtility.HtmlDecode(input);
        }

        return returnValue;
    }

    public DataSet DecodeXml(DataSet dsInput)
    {
        DataSet returnValue = dsInput;
        if (returnValue != null)
        {
            if (returnValue.Tables[0] != null)
            {
                DataTable dtRemove = returnValue.Tables[0];
                DataTable dtAdd = DecodeXml(returnValue.Tables[0]);
                if (returnValue.Tables.CanRemove(dtRemove))
                {
                    returnValue.Tables.Remove(dtRemove);
                    returnValue.Tables.Add(dtAdd);
                }
            }
        }

        return returnValue;
    }
    public DataTable DecodeXml(DataTable dtInput)
    {
        DataTable returnValue = dtInput;
        if (returnValue != null)
        {
            foreach (DataRow dr in returnValue.Rows)
            {
                foreach (DataColumn dataColumn in returnValue.Columns)
                {
                    if (dataColumn.DataType == System.Type.GetType("System.String"))
                    {
                        if (dr[dataColumn] != null)
                        {
                            dr[dataColumn] = WebUtility.HtmlDecode(dr[dataColumn].ToString());
                        }
                        else
                        {
                            dr[dataColumn] = "";
                        }
                    }
                }
            }
        }

        return returnValue;
    }

    public bool LogEmailPending(
        string recipient
        , string sender
        , string replyTo
        , string cc
        , string bcc
        , string subject
        , string emailText
        , string mimeType
        , string emailType
        , string affiliateKey
        , string fromUserId
        , string toUserId
        , string userName
        , SqlConnection myConn
    )
    {
        bool returnValue = false;

        try
        {
            if (myConn.State != ConnectionState.Open)
            {
                myConn.Open();
            }
            String sqlCommand = @"insert into EmailSystem.dbo.Email
                                            (
						                        Recipient
						                        ,Sender
						                        ,ReplyTo
						                        ,Cc
						                        ,Bcc
						                        ,Subject
						                        ,EmailText
						                        ,MimeType
						                        ,EmailType
						                        ,EmailStatus
						                        ,EmailStatusDate
						                        ,AffiliateKey
						                        ,FromUserId
						                        ,ToUserId
						                        ,CreatedByUsername
						                        ,ChangedByUsername
                                            )
                                        values
                                            (
						                        @Recipient
						                        ,@Sender
						                        ,@ReplyTo
						                        ,@Cc
						                        ,@Bcc
						                        ,@Subject
						                        ,@EmailText
						                        ,@MimeType
						                        ,@EmailType
						                        ,'Pending'
						                        ,getdate()						
						                        ,@AffiliateKey
						                        ,@FromUserId
						                        ,@ToUserId
						                        ,@UserName
						                        ,@UserName
                                            );";
            SqlCommand myComm = new SqlCommand(sqlCommand, myConn);
            myComm.CommandTimeout = 6000;
            myComm.Parameters.Add("@Recipient", SqlDbType.NVarChar, -1).Value = recipient;
            myComm.Parameters.Add("@Sender", SqlDbType.NVarChar, -1).Value = sender;
            myComm.Parameters.Add("@ReplyTo", SqlDbType.NVarChar, -1).Value = replyTo;
            myComm.Parameters.Add("@Cc", SqlDbType.NVarChar, -1).Value = cc;
            myComm.Parameters.Add("@Bcc", SqlDbType.NVarChar, -1).Value = bcc;
            myComm.Parameters.Add("@Subject", SqlDbType.NVarChar, -1).Value = subject;
            myComm.Parameters.Add("@EmailText", SqlDbType.NVarChar, -1).Value = emailText;
            myComm.Parameters.Add("@MimeType", SqlDbType.NVarChar, -1).Value = mimeType;
            myComm.Parameters.Add("@EmailType", SqlDbType.NVarChar, -1).Value = emailType;
            myComm.Parameters.Add("@AffiliateKey", SqlDbType.NVarChar, -1).Value = affiliateKey;
            myComm.Parameters.Add("@FromUserId", SqlDbType.NVarChar, -1).Value = fromUserId;
            myComm.Parameters.Add("@ToUserId", SqlDbType.NVarChar, -1).Value = toUserId;
            myComm.Parameters.Add("@UserName", SqlDbType.NVarChar, -1).Value = userName;
            myComm.ExecuteNonQuery();

            returnValue = true;
        }
        catch (Exception e)
        {
            returnValue = false;
            logAction("logEmailPending: EXCEPTION: " + e.ToString());
        }

        return returnValue;
    }

    public void SendMail
    (
        string recipient
        , string sender
        , string replyTo
        , string cc
        , string bcc
        , string subject
        , string emailText
        , List<string> attachments
    )
    {
        string emailAccount = "newsletter@ebooklibrary.org";
        if (!string.IsNullOrEmpty(smtpAccount))
        {
            emailAccount = smtpAccount;
        }
        string emailPassword = "123456#EDC";
        if (!string.IsNullOrEmpty(smtpPassword))
        {
            emailPassword = smtpPassword;
        }
        string emailSMTP = "10.20.10.5";
        if (!string.IsNullOrEmpty(smtpServer))
        {
            emailSMTP = smtpServer;
        }
        int emailSMTPPort = 25;
        if (!string.IsNullOrEmpty(smptPort))
        {
            try
            {
                emailSMTPPort = Convert.ToInt16(smptPort);
            }
            catch { }
        }
        bool isEmailSent = false;
        string smtpSecondary = emailSMTP;

            //  Send through Gmail!
            if (!isEmailSent)
        {
            try
            {
                NetworkCredential loginInfo = new NetworkCredential(emailAccount, emailPassword);

                MailMessage msg = new MailMessage();
                msg.IsBodyHtml = true;
                msg.From = new MailAddress(sender);
                msg.Sender = new MailAddress(sender);
                if (string.IsNullOrEmpty(replyTo))
                {
                    replyTo = sender;
                    if (replyTo.ToLower().Trim() != "support@worldlibrary.org")
                    {
                        msg.ReplyToList.Add(new MailAddress("support@worldlibrary.org"));
                    }
                }
                msg.ReplyToList.Add(new MailAddress(replyTo));

                if (!string.IsNullOrEmpty(recipient))
                {
                    foreach (string emailSemi in recipient.Split(';'))
                    {
                        foreach (string commaEmail in emailSemi.Split(','))
                        {
                            msg.To.Add(new MailAddress(commaEmail));
                        }
                    }
                }
                else
                {
                    msg.To.Add(new MailAddress("support@worldlibrary.org"));
                    msg.To.Add(new MailAddress("jason@worldlibrary.org"));
                }
                if (!string.IsNullOrEmpty(cc))
                {
                    foreach (string emailSemi in cc.Split(';'))
                    {
                        foreach (string commaEmail in emailSemi.Split(','))
                        {
                            msg.CC.Add(new MailAddress(commaEmail));
                        }
                    }
                }
                if (!string.IsNullOrEmpty(bcc))
                {
                    foreach (string emailSemi in bcc.Split(';'))
                    {
                        foreach (string commaEmail in emailSemi.Split(','))
                        {
                            msg.Bcc.Add(new MailAddress(commaEmail));
                        }
                    }
                }
                msg.Subject = subject;
                msg.Body = emailText;
                msg.IsBodyHtml = true;

                try
                {
                    foreach (string attachment in attachments)
                    {
                        if (File.Exists(attachment))
                        {
                            msg.Attachments.Add(new Attachment(attachment));
                        }
                    }
                }
                catch (Exception ex)
                {
                    logAction("SendMail: ERROR: Could not attach files: " + ex.ToString());
                }

                msg.Headers.Add("Precedence", "bulk");
                SmtpClient client = new SmtpClient(smtpSecondary);
                client.Port = emailSMTPPort;

                client.UseDefaultCredentials = false;
                client.Credentials = loginInfo;
                //client.EnableSsl = true;
                client.Send(msg);
                isEmailSent = true;

                logAction("SendMail: INFO: Successfully sent email message");
            }
            catch (Exception ex)
            {
                string errorMessage = "SendMail: EXCEPTION: Could send mesage for.  Error Error is: \r\n" + ex.ToString();
                logAction(errorMessage);
                Console.Write(errorMessage);
            }
        }
        else
        {
            logAction("SendMail: INFO: Email already sent");
        }
    }

    public string CleanDirectoryName(string input)
    {
        string returnValue = input;

        if (!string.IsNullOrEmpty(returnValue))
        {
            returnValue = returnValue.ToLower().Replace(":", "_").Replace("<", "_").Replace(">", "_").Replace("\"", "_").Replace("|", "_").Replace("?", "_").Replace("*", "_").Replace(@"con.", @"con_")
.Replace(@"prn.", @"prn_")
.Replace(@"aux.", @"aux_")
.Replace(@"nul.", @"nul_")
.Replace(@"com1.", @"com1_")
.Replace(@"com2.", @"com2_")
.Replace(@"com3.", @"com3_")
.Replace(@"com4.", @"com4_")
.Replace(@"com5.", @"com5_")
.Replace(@"com6.", @"com6_")
.Replace(@"com7.", @"com7_")
.Replace(@"com8.", @"com8_")
.Replace(@"com9.", @"com9_")
.Replace(@"lpt1.", @"lpt1_")
.Replace(@"lpt2.", @"lpt2_")
.Replace(@"lpt3.", @"lpt3_")
.Replace(@"lpt4.", @"lpt4_")
.Replace(@"lpt5.", @"lpt5_")
.Replace(@"lpt6.", @"lpt6_")
.Replace(@"lpt7.", @"lpt7_")
.Replace(@"lpt8.", @"lpt8_")
.Replace(@"lpt9.", @"lpt9_");
        }

        return returnValue;
    }


    /// <summary>
    /// Encryption section  "2RleqdHjMLseKHFzRNq4Bg****"
    /// </summary>
    // This constant string is used as a "salt" value for the PasswordDeriveBytes function calls.
    // This size of the IV (in bytes) must = (keysize / 8).  Default keysize is 256, so the IV must be
    // 32 bytes long.  Using a 16 character string here gives us 32 bytes when converted to a byte array.
    private static readonly byte[] initVectorBytes = Encoding.ASCII.GetBytes("1234567890zxcvbn");
    private static readonly string standardPassPhrase = "honteux -- abashes : penaud -- abashment";

    // This constant is used to determine the keysize of the encryption algorithm.
    private const int keysize = 256;

    public string Encrypt(string plainText)
    {
        return Encrypt(plainText, standardPassPhrase).Replace("/", "~~").Replace("+", "^^").Replace("=", "**");
    }
    public string Encrypt(string plainText, string passPhrase)
    {
        byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);
        using (PasswordDeriveBytes password = new PasswordDeriveBytes(passPhrase, null))
        {
            byte[] keyBytes = password.GetBytes(keysize / 8);
            using (RijndaelManaged symmetricKey = new RijndaelManaged())
            {
                symmetricKey.Mode = CipherMode.CBC;
                using (ICryptoTransform encryptor = symmetricKey.CreateEncryptor(keyBytes, initVectorBytes))
                {
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        using (CryptoStream cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
                        {
                            cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
                            cryptoStream.FlushFinalBlock();
                            byte[] cipherTextBytes = memoryStream.ToArray();
                            return Convert.ToBase64String(cipherTextBytes);
                        }
                    }
                }
            }
        }
    }

    public string Decrypt(string cipherText)
    {
        return Decrypt(cipherText.Replace("~~", "/").Replace("^^", "+").Replace("**", "="), standardPassPhrase);
    }
    public string Decrypt(string cipherText, string passPhrase)
    {
        cipherText = cipherText.Replace(' ', '+');
        byte[] cipherTextBytes = Convert.FromBase64String(cipherText);
        using (PasswordDeriveBytes password = new PasswordDeriveBytes(passPhrase, null))
        {
            byte[] keyBytes = password.GetBytes(keysize / 8);
            using (RijndaelManaged symmetricKey = new RijndaelManaged())
            {
                symmetricKey.Mode = CipherMode.CBC;
                using (ICryptoTransform decryptor = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes))
                {
                    using (MemoryStream memoryStream = new MemoryStream(cipherTextBytes))
                    {
                        using (CryptoStream cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read))
                        {
                            byte[] plainTextBytes = new byte[cipherTextBytes.Length];
                            int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
                            return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount);
                        }
                    }
                }
            }
        }
    }

    public string GetMd5HashForString(string input, string version = "2")
    {
        string returnValue = "";

        if (version == "2")
        {
            // Best Hashing alorithm, but not the one wikipedia uses all the time...
            byte[] encoded = new System.Text.UTF8Encoding().GetBytes(input);
            byte[] hash = ((HashAlgorithm)CryptoConfig.CreateFromName("MD5")).ComputeHash(encoded);
            returnValue = BitConverter.ToString(hash).Replace("-", string.Empty).ToLower();
        }

        if (version == "1")
        {
            // Results are not fully UTF-8 but produce the same hash values as wikipedia
            // byte array representation of that string
            byte[] asciiBytes = ASCIIEncoding.Default.GetBytes(input);
            // need MD5 to calculate the hash
            byte[] hashedBytes = MD5CryptoServiceProvider.Create().ComputeHash(asciiBytes);
            // string representation (similar to UNIX format), without dashes, in lowercase
            returnValue = BitConverter.ToString(hashedBytes).Replace("-", "").ToLower();
        }

        return returnValue;
    }
}

public static class HttpWebResponseExt
{
    public static HttpWebResponse GetResponseNoException(this HttpWebRequest req)
    {
        try
        {
            return (HttpWebResponse)req.GetResponse();
        }
        catch (WebException we)
        {
            var resp = we.Response as HttpWebResponse;
            if (resp == null)
                throw;
            return resp;
        }
    }
}

class UserProfile
{
    private string userAgent = "";
    public string UserAgent
    {
        get { return userAgent; }
        set { userAgent = value; }
    }
    private string referer = "https://www.google.com/";
    public string Referer
    {
        get { return referer; }
        set { referer = value; }
    }
    private string lastPage = "";
    public string LastPage
    {
        get { return lastPage; }
        set { lastPage = value; }
    }
    private string method = "GET";
    public string Method
    {
        get { return method; }
        set { method = value; }
    }
    private CookieContainer cookieJar = new CookieContainer();
    public CookieContainer CookieJar
    {
        get { return cookieJar; }
        set { cookieJar = value; }
    }

    public UserProfile()
    {
        ResetUserProfile(false);
    }
    public UserProfile(string inputUserAgent)
    {
        userAgent = inputUserAgent;
    }

    public void ResetUserProfile(bool isKeepUserAgent = true)
    {
        if (!isKeepUserAgent)
        {
            userAgent = GetRandomUserAgent();
        }
        cookieJar = new CookieContainer();
        referer = "";
        lastPage = "";
    }

    public string GetRandomUserAgent(bool isTryAgain = true)
    {
        string returnValue = "";

        try
        {
            Random rand = new Random();
            returnValue = arrUserAgents[rand.Next(0, arrUserAgents.Length)];
        }
        catch { }
        if (string.IsNullOrEmpty(returnValue) && isTryAgain)
        {
            returnValue = GetRandomUserAgent(false);
        }

        return returnValue;
    }
    string[] arrUserAgents =
        {
            @"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36"
            ,@"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36"
            ,@"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_4) AppleWebKit/601.5.17 (KHTML, like Gecko) Version/9.1 Safari/601.5.17"
            ,@"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.94 Safari/537.36"
            ,@"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.94 Safari/537.36"
            ,@"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.94 Safari/537.36"
            ,@"Mozilla/5.0 (Windows NT 6.1; WOW64; rv:46.0) Gecko/20100101 Firefox/46.0"
            ,@"Mozilla/5.0 (Windows NT 10.0; WOW64; rv:46.0) Gecko/20100101 Firefox/46.0"
            ,@"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36"
            ,@"Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/49.0.2623.112 Safari/537.36"
            ,@"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/49.0.2623.112 Safari/537.36"
            ,@"Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko"
            ,@"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_4) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.86 Safari/537.36"
            ,@"Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36"
            ,@"Mozilla/5.0 (Macintosh; Intel Mac OS X 10.11; rv:46.0) Gecko/20100101 Firefox/46.0"
            ,@"Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.94 Safari/537.36"
            ,@"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36"
            ,@"Mozilla/5.0 (Windows NT 6.1; WOW64; rv:45.0) Gecko/20100101 Firefox/45.0"
            ,@"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/601.6.17 (KHTML, like Gecko) Version/9.1.1 Safari/601.6.17"
            ,@"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.94 Safari/537.36"
            ,@"Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36"
        };
}
}