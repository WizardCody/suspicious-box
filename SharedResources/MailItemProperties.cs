using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Outlook = Microsoft.Office.Interop.Outlook;

namespace SharedResources
{
    class MailItemProperties
    {
        /// <summary>
        /// Canonical name: PidTagTransportMessageHeaders
        /// Description: Contains transport-specific message envelope information for email.
        /// Property ID:0x007D
        /// Data type:PtypString, 0x001F
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static string GetHeader(Outlook.MailItem item)
        {
            try
            {
                return item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F");
            }
            catch (Exception exc)
            {
                return exc.Message;
            }
        }

        /// <summary>
        /// Canonical name: PidTagSmtpAddress
        /// it can be Recipient type or AddressEntry type.
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static string GetSMTPAddressForRecipient(Outlook.MailItem item)
        {
            try
            {
                return item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E");
            }
            catch (Exception exc)
            {
                return exc.Message;
            }
        }

        /// <summary>
        /// Canonical name: PidTagLastVerbExecuted
        /// Description: Specifies the last verb executed for the message item to which it is related.
        /// Property ID: 0x1081
        /// Data type: PtypInteger32, 0x0003
        ///
        /// Last_Verb_Reply_All = 103
        /// Last_Verb_Reply_Sender = 102
        /// Last_Verb_Reply_Forward = 104
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static string GetReplied(Outlook.MailItem item)
        {
            try
            {
                var result = item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003");

                if (result == 103)
                    return ("Reply_All");
                else if (result == 102)
                    return ("Reply_Sender");
                else if (result == 104)
                    return ("Reply_Forward");
                else
                    return ("No");
            }
            catch (Exception exc)
            {
                return exc.Message;
            }
        }

        /// <summary>
        /// Canonical name: PidTagLastVerbExecutionTime
        /// Description: Contains the date and time, in UTC, during which the operation represented in the PidTagLastVerbExecuted property took place.
        /// Property ID: 0x1082
        /// Data type: PtypTime, 0x0040
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static string GetRepliedTime(Outlook.MailItem item)
        {
            try
            {
                // System.DateTime
                return item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10820040").ToString();
            }
            catch (Exception exc)
            {
                return exc.Message;
            }
        }

        public static bool GetIsReply(Outlook.MailItem item)
        {
            bool result = false;

            if (item.ConversationIndex.Length > 44)
            {
                result = true;
            }

            return result;
        }
        

    }
}
