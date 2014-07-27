using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics.Tracing;

namespace RptToXml
{
    [EventSource(Name = "RptToXML")]
    public class RptToXMLEventSource : EventSource
    {

        #region Inner Types

        public class Keywords
        {
            public const EventKeywords Diagnostic = (EventKeywords)1;
            public const EventKeywords Perf = (EventKeywords)2;
            public const EventKeywords XMLProcessing = (EventKeywords)4;
            //public const EventKeywords XMLAttributeProcessing = (EventKeywords)8;
        }

        public class Tasks
        {
            public const EventTask Startup = (EventTask)1;
            public const EventTask Shutdown = (EventTask)2;
        }

        #endregion Inner Types


        # region Private Fields

        private static readonly Lazy<RptToXMLEventSource> Instance =
            new Lazy<RptToXMLEventSource>(() => new RptToXMLEventSource());

        # endregion Private Fields

        #region Private Constructor

        private RptToXMLEventSource() { }
 
        #endregion Private Constructor

        #region Public Properties
        
        public static RptToXMLEventSource Log { get { return Instance.Value; } }
        
        #endregion Public Properties

        #region Non-Event Public Methods
        
        #endregion Non-Event Public Methods

        #region Event Public methods

        [Event(
            1,
            Message = "Argument Error: {0}", 
            Level = EventLevel.Error,
            Keywords = Keywords.Diagnostic)]
        public void ArgumentError(string message)
        {
            if (this.IsEnabled())
            {
                WriteEvent(1, message);
            }
        }

        [Event(
            2,
            Message = "Application Information: {0}",
            Keywords = Keywords.Diagnostic,
            Level = EventLevel.Verbose)]
        public void Info(string message)
        {
            if (this.IsEnabled())
            {
                WriteEvent(2, message);
            }
        }

        [Event(
            3,
            Message = "Application Warning: {0}",
            Level = EventLevel.Warning,
            Keywords = Keywords.Diagnostic)]
        public void Warning(string message)
        {
            if (this.IsEnabled())
            {
                WriteEvent(3, message);
            }
        }

        [Event(
            4,
            Message = "XML Processing: {0}",
            Level = EventLevel.Verbose,
            Keywords = Keywords.XMLProcessing)]
                public void XMLProcessing(string message)
                {
                    if (this.IsEnabled())
                    {
                        WriteEvent(4, message);
                    }
                }

        [Event(
            5,
            Message = "Unhandled ConditionFormula: {0}",
            Level = EventLevel.Error)]
                public void UnhandledConditionFormula(string message)
                {
                    if (this.IsEnabled())
                    {
                        WriteEvent(5, message);
                    }
                }

        [Event(
            6,
            Message = "Starting RptToXML",
            Level = EventLevel.Informational)]
        public void Startup()
        {
                WriteEvent(6);
                Console.WriteLine("Starting xml dump.");
        }

        [Event(
            7,
            Message = "Finishing RptToXML",
            Level = EventLevel.Informational)]
        public void Shutdown()
        {
                WriteEvent(7);
                Console.WriteLine("Finishing xml dump.");
        }

        #endregion Event Public methods

    }
}
