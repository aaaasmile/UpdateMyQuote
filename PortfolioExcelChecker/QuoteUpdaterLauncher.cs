using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace PortfolioExcelChecker
{
    public class QuoteUpdaterLauncher
    {

        private StringBuilder _result = new StringBuilder();

        public delegate void ProcessTerminatedHandler(string Result);
        public event ProcessTerminatedHandler TeminatedEvent;


        public QuoteUpdaterLauncher()
        {
        }
        public void StartProcess(string fileName)
        {

            ProcessStarter processStarter = new ProcessStarter();
            processStarter.OutputWrittenEvent += (x) => _result.AppendLine(x);

            string rubyExePath = GetRubyExePath();
            processStarter.ExecuteCmd(rubyExePath, fileName);

            FireTeminatedEvent();
        }

        public void CheckVersion()
        {
            ProcessStarter processStarter = new ProcessStarter();
            processStarter.OutputWrittenEvent += (x) => _result.AppendLine(x);

            string rubyExePath = GetRubyExePath();
            try
            {
                processStarter.ExecuteCmd(rubyExePath, "-v");
            }
            catch (Exception ex)
            {

                _result.Append(string.Format("Error: {0}", ex));
            }


            FireTeminatedEvent();
        }

        private string GetRubyExePath()
        {
            return @"D:\ruby\ruby_2_3_1\bin\ruby.exe";
        }

       

        private void FireTeminatedEvent()
        {
            if (TeminatedEvent != null)
            {
                TeminatedEvent(_result.ToString());
            }
        }

    }
}
