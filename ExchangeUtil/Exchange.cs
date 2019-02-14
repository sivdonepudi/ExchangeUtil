using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeUtil
{
    public static class ExchangeHelper
    {
        public static string Login { get; set; }

        public static string Password { get; set; }

        public static string Url { get; set; }

        public static int BackDays { get; set; }

        public static string User { get; set; }

        private static ExchangeService _globalService;

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public static ObservableCollection<ExchangeData> GetUserAppointments()
        {
            ObservableCollection<ExchangeData> lstData = new ObservableCollection<ExchangeData>();

            try
            {
                if (_globalService == null)
                    _globalService = GetExchangeServiceObject(new TraceListner());

                if (!_globalService.HttpHeaders.ContainsKey("X-AnchorMailbox"))
                {
                    _globalService.HttpHeaders.Add("X-AnchorMailbox", User);
                }
                else
                {
                    _globalService.HttpHeaders["X-AnchorMailbox"] = User;
                }

                _globalService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, User);
                // AppointmentSchema.Id
                PropertySet basicProps = new PropertySet(BasePropertySet.IdOnly,
                    AppointmentSchema.Start,
                    AppointmentSchema.ICalUid,
                    ItemSchema.Subject,
                    ItemSchema.Id,
                    AppointmentSchema.Organizer,
                    ItemSchema.Categories);

                DateTime dtStart = DateTime.Now.AddDays(BackDays);

                CalendarView calView = new CalendarView(dtStart, dtStart.AddYears(2))
                {
                    Traversal = ItemTraversal.Shallow,
                    PropertySet = new PropertySet(BasePropertySet.IdOnly, AppointmentSchema.Start),
                    MaxItemsReturned = 500
                };


                List<Appointment> appointments = new List<Appointment>();
                FindItemsResults<Appointment> userAppointments =
                    _globalService.FindAppointments(WellKnownFolderName.Calendar, calView);
                if ((userAppointments != null) && (userAppointments.Any()))
                {
                    appointments.AddRange(userAppointments);
                    while (userAppointments.MoreAvailable)
                    {
                        calView.StartDate = appointments.Last().Start;
                        userAppointments = _globalService.FindAppointments(WellKnownFolderName.Calendar, calView);
                        appointments.AddRange(userAppointments);
                    }

                    ServiceResponseCollection<ServiceResponse> response =
                        _globalService.LoadPropertiesForItems(appointments, basicProps);
                    if (response.OverallResult != ServiceResult.Success) // property load not success
                    {
                        return null;
                    }

                    if (appointments?.Count > 0)
                    {
                        foreach (Appointment item in appointments)
                        {
                            lstData.Add(new ExchangeData
                            {
                                UniqueId = item.Id.UniqueId,
                                StartDate = item.Start.ToString("dd-MM-yyyy HH:mm"),
                                Subject = item.Subject,
                                Organizer = item.Organizer.Address,
                                Categories = item.Categories.ToString(),
                                IsSelected = (item.Subject.StartsWith("Canceled:", StringComparison.OrdinalIgnoreCase)
                                              || item.Subject.StartsWith("Abgesagt:", StringComparison.OrdinalIgnoreCase))
                            });
                        }

                        //group by with Subject, start date
                        var duplicates = from p in lstData.Where(p => !p.IsSelected)
                                         group p by new { p.Subject, p.StartDate }
                                            into q
                                         where q.Count() > 1
                                         select new ExchangeData()
                                         {
                                             Subject = q.Key.Subject,
                                             StartDate = q.Key.StartDate
                                         };
                        var exchangeDatas = duplicates as ExchangeData[] ?? duplicates.ToArray();
                        if (exchangeDatas?.Length > 0)
                        {
                            foreach (var item in lstData.Where(p => !p.IsSelected))
                            {
                                if (exchangeDatas.Any(p => string.Equals(p.Subject, item.Subject)
                                                        && string.Equals(p.StartDate, item.StartDate)))
                                    item.IsSelected = true;
                            }
                        }
                    }

                }
            }
            catch (Exception e)
            {
                WriteLog(e.Message);
            }

            return lstData;
        }


        private static bool CertificateValidationCallback(object sender,
            X509Certificate certificate,
            X509Chain chain,
            SslPolicyErrors sslPolicyErrors)
        {
            // If the certificate is a valid, signed certificate, return true.
            if (sslPolicyErrors == SslPolicyErrors.None)
            {
                return true;
            }

            // If there are errors in the certificate chain, look at each error to determine the cause.
            if ((sslPolicyErrors & SslPolicyErrors.RemoteCertificateChainErrors) != 0)
            {
                if ((chain != null) && (chain.ChainStatus != null))
                {
                    foreach (X509ChainStatus status in chain.ChainStatus)
                    {
                        if ((certificate.Subject == certificate.Issuer) &&
                            (status.Status == X509ChainStatusFlags.UntrustedRoot))
                        {
                            // Self-signed certificates with an untrusted root are valid. 
                        }
                        else
                        {
                            if (status.Status != X509ChainStatusFlags.NoError)
                            {
                                // If there are any other errors in the certificate chain, the certificate is invalid,
                                // so the method returns false.
                                return false;
                            }
                        }
                    }
                }

                // When processing reaches this line, the only errors in the certificate chain are 
                // untrusted root errors for self-signed certificates. These certificates are valid
                // for default Exchange server installations, so return true.
                return true;
            }

            // In all other cases, return false.
            return false;
        }

        private static ExchangeService GetExchangeServiceObject(ITraceListener listner)
        {
            string methodName = "GetExchangeServiceObject";
            try
            {
                ServicePointManager.ServerCertificateValidationCallback = CertificateValidationCallback;
                ExchangeService service = new ExchangeService
                {
                    Credentials = new WebCredentials(Login, Password),
                    Url = new Uri(Url),
                    PreferredCulture = GetCulture()
                };
                if (listner != null)
                {
                    service.TraceEnabled = true;
                    service.TraceFlags = TraceFlags.All;
                    service.TraceListener = listner;
                }

                return service;
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }

            return null;
        }

        private static CultureInfo GetCulture()
        {
            string methodName = "GetCulture";
            string strCulture = "de-DE";
            CultureInfo culInfo = new CultureInfo(strCulture);
            try
            {
                culInfo = new CultureInfo(strCulture);
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }

            return culInfo;
        }

        private static readonly object LogObject = new object();

        private static void WriteLog(string message)
        {
            lock (LogObject)
            {
                string strLogFilePath = AppDomain.CurrentDomain.BaseDirectory + Path.DirectorySeparatorChar +
                                        "ExchangeUtil.log";
                StringBuilder sbMessage = new StringBuilder();
                using (StreamWriter sw = new StreamWriter(strLogFilePath, true))
                {
                    sbMessage.Append("\n ");
                    sbMessage.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    sbMessage.Append(" : " + message);
                    sw.WriteLine(sbMessage.ToString());
                    sw.Flush();
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public class ExchangeData
        {
            public string UniqueId { get; set; }

            public string Subject { get; set; }

            public string StartDate { get; set; }

            public string Organizer { get; set; }

            public bool IsSelected { get; set; }

            public string Categories { get; set; }
        }

        public static void DeleteItems(List<string> lstDeleteItems)
        {
            try
            {


                if (_globalService == null)
                    _globalService = GetExchangeServiceObject(new TraceListner());

                if (!_globalService.HttpHeaders.ContainsKey("X-AnchorMailbox"))
                {
                    _globalService.HttpHeaders.Add("X-AnchorMailbox", User);
                }
                else
                {
                    _globalService.HttpHeaders["X-AnchorMailbox"] = User;
                }

                _globalService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, User);

                Collection<ItemId> itemIds = new Collection<ItemId>();
                List<int> lstNotFound = new List<int>();
                foreach (string strId in lstDeleteItems)
                {
                    try
                    {
                        Item _item = Item.Bind(_globalService, new ItemId(strId), PropertySet.IdOnly);
                        if (_item != null)
                        {
                            itemIds.Add(_item.Id);
                        }
                    }
                    catch (ServiceResponseException ex)
                    {
                        if (ex.ErrorCode == ServiceError.ErrorItemNotFound) // item not found 
                        {
                            WriteLog("Item is not found " + strId);
                        }
                        else
                        {
                            WriteLog("Error  " + ex.ErrorCode.ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteLog(ex.Message);
                    }
                }

                IEnumerable<IEnumerable<ItemId>> batchedList = itemIds.Batch(500);
                int count = 0;
                if (batchedList.Any())
                {
                    StringBuilder sbErrorMsgs = new StringBuilder();
                    List<ServiceError> lstError = new List<ServiceError>();
                    foreach (IEnumerable<ItemId> batch_Items in batchedList)
                    {
                        var batchItems = batch_Items as ItemId[] ?? batch_Items.ToArray();
                        if (batchItems.Any())
                        {
                            ServiceResponseCollection<ServiceResponse> responses = _globalService.DeleteItems(
                                batchItems, DeleteMode.HardDelete,
                                SendCancellationsMode.SendToNone,
                                AffectedTaskOccurrence.AllOccurrences);
                            if (responses.OverallResult == ServiceResult.Success)
                            {
                                count++;
                            }

                            foreach (ServiceResponse resp in responses)
                            {
                                if (resp.Result != ServiceResult.Success)
                                {
                                    if (!lstError.Contains(resp.ErrorCode))
                                    {
                                        sbErrorMsgs.Append(string.Format("\r\n{0}: {1}", "ResultText", resp.Result));
                                        sbErrorMsgs.Append(string.Format("\r\n{0}: {1}", "ErrorCodeText",
                                            resp.ErrorCode));
                                        sbErrorMsgs.Append(string.Format("\r\n{0}: {1}", "ErrorMessageText",
                                            resp.ErrorMessage));
                                        lstError.Add(resp.ErrorCode);
                                    }
                                }
                            }
                        }
                    }

                    if (!string.IsNullOrEmpty(sbErrorMsgs.ToString()))
                    {
                        WriteLog(sbErrorMsgs.ToString());
                    }
                }
            }

            catch (Exception e)
            {
                WriteLog(e.Message);
            }
        }
    }


    public static class BatchLinq
    {
        public static IEnumerable<IEnumerable<TSource>> Batch<TSource>(
            this IEnumerable<TSource> source, int size)
        {
            TSource[] bucket = null;
            var count = 0;

            foreach (var item in source)
            {
                if (bucket == null)
                {
                    bucket = new TSource[size];
                }

                bucket[count++] = item;
                if (count != size)
                {
                    continue;
                }

                yield return bucket;

                bucket = null;
                count = 0;
            }

            if ((bucket != null) && (count > 0))
            {
                yield return bucket.Take(count);
            }
        }

        public static string Replace(this string source, string oldValue, string newValue, StringComparison comparisonType)
        {
            if ((source.Length == 0) || (oldValue.Length == 0))
            {
                return source;
            }

            var result = new StringBuilder();
            int startingPos = 0;
            int nextMatch;
            while ((nextMatch = source.IndexOf(oldValue, startingPos, comparisonType)) > -1)
            {
                result.Append(source, startingPos, nextMatch - startingPos);
                result.Append(newValue);
                startingPos = nextMatch + oldValue.Length;
            }

            result.Append(source, startingPos, source.Length - startingPos);

            return result.ToString();
        }
    }

    public class TraceListner : ITraceListener
    {
        /// <summary>
        /// Interface method
        /// </summary>
        /// <param name="traceType"></param>
        /// <param name="traceMessage"></param>
        public void Trace(string traceType, string traceMessage)
        {
            CreateXmlTextFile(traceType, traceMessage);
        }

        /// <summary>
        /// Creates XML text file
        /// </summary>
        /// <param name="fileName">contains file name</param>
        /// <param name="traceContent">contains trace content</param>
        private static void CreateXmlTextFile(string fileName, string traceContent)
        {
            string strTraceFilePath = AppDomain.CurrentDomain.BaseDirectory + Path.DirectorySeparatorChar + (fileName);
            try
            {
                XElement xDoc = XElement.Parse(traceContent);
                xDoc.Save(strTraceFilePath + ".xml");
            }
            catch
            {
                File.WriteAllText(strTraceFilePath + ".txt", traceContent);
            }
        }
    }

}
