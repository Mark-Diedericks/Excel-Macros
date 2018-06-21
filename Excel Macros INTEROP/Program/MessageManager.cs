/*
 * Mark Diedericks
 * 09/06/2015
 * Version 1.0.0
 * Handling message events (Message Boxes) between backend and UI
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_INTEROP
{
    public class MessageManager
    {
        //VoidMessage event, for all Forms and GUIs
        public delegate void VoidMessageEvent(string content, string title);
        public event VoidMessageEvent DisplayOkMessageEvent;

        //ObjectMessage event, for all Forms and GUIs
        public delegate void ObjectMessageEvent(string content, string title, Action<bool> OnReturn);
        public event ObjectMessageEvent DisplayYesNoMessageEvent;

        //InputMessage event, for all Forms and GUIs
        public delegate void InputMessageEvent(string message, object title, object def, object left, object top, object helpFile, object helpContextID, object type, Action<object> OnResult);
        public event InputMessageEvent DisplayInputMessageEvent;

        private static MessageManager s_Instance;

        private MessageManager()
        {
            s_Instance = this;
        }

        public static void Instantiate()
        {
            new MessageManager();
        }

        public static MessageManager GetInstance()
        {
            return s_Instance != null ? s_Instance : new MessageManager();
        }

        public static void DisplayOkMessage(string content, string title)
        {
            GetInstance().DisplayOkMessageEvent?.Invoke(content, title);
        }

        public static void DisplayYesNoMessage(string content, string title, Action<bool> OnReturn)
        {
            GetInstance().DisplayYesNoMessageEvent?.Invoke(content, title, OnReturn);
        }

        public static void DisplayInputMessage(string message, object title, object def, object left, object top, object helpFile, object helpContextID, object type, Action<object> OnResult)
        {
            GetInstance().DisplayInputMessageEvent?.Invoke(message, title, def, left, top, helpFile, helpContextID, type, OnResult);
        }

    }
}
