using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SimpleConverter.Contract;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    public interface IMessenger
    {
        void SendMessage(string message, MessageLevel level = MessageLevel.INFO);
    }
}
