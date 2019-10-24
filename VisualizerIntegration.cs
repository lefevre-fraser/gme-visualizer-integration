using System;
using System.Collections.Generic;
using System.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.IO;
using System.IO.Pipes;
using GME.CSharp;
using GME.MGA;
using GME;
using GME.MGA.Core;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading;

namespace VisualizerIntegration
{
    //[ComVisible(true)]
    //public class GMEConsoleReference
    //{
    //    public GMEConsole GMEConsole = null;
    //}

    [Guid(ComponentConfig.guid),
    ProgId(ComponentConfig.progID),
    ClassInterface(ClassInterfaceType.AutoDual)]
    [ComVisible(true)]
    public class VisualizerIntegrationAddon : IMgaComponentEx, IGMEVersionInfo, IMgaEventSink/*, IDisposable*/
    {
        private MgaAddOn addon;
        private bool componentEnabled = true;
        private bool handleEvents = true;
        private MgaProject project = null;
        private List<string> openModels = new List<string>();
        private const int BUFFER_SIZE = 256;
        private static Thread thread = null;
        private static string mgaFile = null;
        private static List<string> modelsToReOpen = new List<string>();
        private static string pipeFile = null;
        private static NamedPipeServerStream namedPipeServer = null;
        private static NamedPipeClientStream namedPipeClient = null;
        GMEConsole GMEConsole { get; set; }
        //private static GMEConsoleReference consoleReference = null;
        private static System.Windows.Forms.Control hiddenWindow;

        public enum Actions
        {
            CLOSE,
            OPEN,
        }

        // Event handlers for addons
        #region MgaEventSink members
        public void GlobalEvent(globalevent_enum @event)
        {
            if (GMEConsole == null || GMEConsole.gme == null)
            {
                GMEConsole = GMEConsole.CreateFromProject(this.project);
                //if (consoleReference.GMEConsole == null)
                //{
                //    consoleReference.GMEConsole = GMEConsole;
                //    if (consoleReference.GMEConsole.gme == null)
                //    {
                //        consoleReference.GMEConsole.gme = GMEConsole.gme;
                //    }
                //}
            }

            if (@event == globalevent_enum.GLOBALEVENT_SAVE_PROJECT)
            {
                // update the list of models to be re-opened
                modelsToReOpen = new List<string>(openModels);

#if(DEBUG)
                // notify open models on console
                GMEConsole.Error.Write(String.Format("Models Open: {0}", openModels.Count));
                foreach(string model in openModels)
                {
                    GMEConsole.Out.Write(model);
                }
#endif
            }

            if (@event == globalevent_enum.GLOBALEVENT_CLOSE_PROJECT)
            {
#if(DEBUG)
                // notify close project
                if (GMEConsole != null)
                {
                    GMEConsole.Error.WriteLine("Closed project: {0}", project.ProjectConnStr.Split('\\').Last());
                }
#endif

                // properly destroy this object
                //if (GMEConsole != null)
                //{
                //    if (GMEConsole.gme != null)
                //    {
                //        Marshal.FinalReleaseComObject(GMEConsole.gme);
                //    }
                //    GMEConsole = null;
                //}
                addon.Destroy();
                Marshal.FinalReleaseComObject(addon);
                addon = null;
                //hiddenWindow.BeginInvoke((System.Action)delegate
                //{
                //    hiddenWindow.Dispose();
                //    hiddenWindow = null;
                //});
            }

            if (@event == globalevent_enum.GLOBALEVENT_OPEN_PROJECT_FINISHED)
            {
#if(DEBUG)
                // notify open project
                if (GMEConsole != null)
                {
                    GMEConsole.Out.WriteLine("Opened project: {0}", project.ProjectConnStr.Split('\\').Last());
                }
#endif

                // Re-opening a project, re-open the previously open models
                if (project != null && mgaFile == project.ProjectConnStr)
                {
                    var modelsToReOpenFcos = modelsToReOpen.Select(path => project.ObjectByPath[path]).Where(fco => fco != null).ToList();
                    if (modelsToReOpenFcos.Count > 0)
                    {
                        hiddenWindow.BeginInvoke((System.Action)delegate
                        {
                            foreach (MgaObject obj in modelsToReOpenFcos)
                            {
                                try
                                {
                                    GMEConsole.gme.ShowFCO((MgaFCO)obj, false);
                                }
                                catch (Exception e)
                                {
                                    GMEConsole.Error.Write(e.ToString());
                                }
                            }
                        });
                    }
                }
                modelsToReOpen.Clear();

                // update the open mga file name
                mgaFile = GMEConsole.gme.MgaProject.ProjectConnStr;

                if (thread != null)
                {
                    thread.Abort();
                    if (namedPipeServer != null && namedPipeServer.IsConnected)
                    {
                        namedPipeClient = new NamedPipeClientStream(".", pipeFile, PipeDirection.InOut);
                        namedPipeClient.Connect();
                        namedPipeClient.Close();
                        namedPipeClient = null;
                    }
                    thread.Join();
                }
                thread = new Thread(() =>
                {
                    try
                    {
                        pipeFile = String.Format("\\{0}\\{1}", Process.GetCurrentProcess().Id, mgaFile.Split('\\').Last());
                        PipeSecurity ps = new PipeSecurity();
                        ps.AddAccessRule(new PipeAccessRule(@"NT AUTHORITY\Everyone", PipeAccessRights.ReadWrite, System.Security.AccessControl.AccessControlType.Allow));
                        namedPipeServer = new NamedPipeServerStream(pipeFile, PipeDirection.InOut, 1, PipeTransmissionMode.Message, PipeOptions.None, BUFFER_SIZE, BUFFER_SIZE, ps);

#if(DEBUG)
                        GMEConsole.Out.Write(pipeFile);
                        //consoleReference.GMEConsole.Out.Write(pipeFile);
#endif
                        namedPipeServer.WaitForConnection();
                        byte[] buffer = new byte[BUFFER_SIZE];
                        int nread = namedPipeServer.Read(buffer, 0, BUFFER_SIZE);
                        string line = Encoding.UTF8.GetString(buffer, 0, nread);
                        Actions act = (Actions)Enum.Parse(typeof(Actions), line, true);
                        if (act != Actions.CLOSE) throw new Exception(String.Format("Expected action: CLOSE, received action: {0}", act));

                        GMEConsole.gme.SaveProject();
                        //consoleReference.GMEConsole.gme.SaveProject();
                        GMEConsole.gme.CloseProject(false);
                        //consoleReference.GMEConsole.gme.CloseProject(false);
                        using (BinaryWriter writer = new BinaryWriter(new MemoryStream()))
                        {
                            writer.Write("closed");
                            namedPipeServer.Write(((MemoryStream)writer.BaseStream).ToArray(), 0, ((MemoryStream)writer.BaseStream).ToArray().Length);
                        }

                        nread = namedPipeServer.Read(buffer, 0, BUFFER_SIZE);
                        line = Encoding.UTF8.GetString(buffer, 0, nread);
                        act = (Actions)Enum.Parse(typeof(Actions), line, true);
                        if (act != Actions.OPEN) throw new Exception(String.Format("Expected action: OPEN, received action: {0}", act));

                        //hiddenWindow.Invoke((Action)delegate
                        //{
                        //    GMEConsole = GMEConsole.CreateFromProject(project);
                        GMEConsole.gme.OpenProject(mgaFile);
                        //consoleReference.GMEConsole.gme.OpenProject(mgaFile);
                        //});
                    }
                    catch (ThreadAbortException e)
                    {
                        
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(String.Format("Inner Exception: {0}\nMessage: {1}\nSource: {2}\nStack Trace: {3}", e.InnerException, e.Message, e.Source, e.StackTrace));
                        //consoleReference.GMEConsole.Error.WriteLine(String.Format("Inner Exception: {0}\nMessage: {1}\nSource: {2}\nStack Trace: {3}", e.InnerException, e.Message, e.Source, e.StackTrace));
                    }
                    finally
                    {
                        if (namedPipeServer.IsConnected)
                        {
                            namedPipeServer.Disconnect();
                        }
                        namedPipeServer.Close();
                        namedPipeServer = null;
                    }
                });
                thread.Start();

                //if (task != null && task.Status.Equals(TaskStatus.Running))
                //{
                    //thread.Abort();
                    //task = null;
                //}
                //task = Task.Factory.StartNew(action, Process.GetCurrentProcess().Id);
            }

            #region Other EventHandlers
            if (@event == globalevent_enum.APPEVENT_XML_IMPORT_BEGIN)
            {
                handleEvents = false;
                addon.EventMask = 0;
            }
            else if (@event == globalevent_enum.APPEVENT_XML_IMPORT_END)
            {
                unchecked { addon.EventMask = (uint)ComponentConfig.eventMask; }
                handleEvents = true;
            }
            else if (@event == globalevent_enum.APPEVENT_LIB_ATTACH_BEGIN)
            {
                addon.EventMask = 0;
                handleEvents = false;
            }
            else if (@event == globalevent_enum.APPEVENT_LIB_ATTACH_END)
            {
                unchecked { addon.EventMask = (uint)ComponentConfig.eventMask; }
                handleEvents = true;
            }
            #endregion

            if (!componentEnabled)
            {
                return;
            }

            // TODO: Handle global events
            // MessageBox.Show(@event.ToString());
        }

        /// <summary>
        /// Called when an FCO or folder changes
        /// </summary>
        /// <param name="subject">the object the event(s) happened to</param>
        /// <param name="eventMask">objectevent_enum values ORed together</param>
        /// <param name="param">extra information provided for cetertain event types</param>
        public void ObjectEvent(MgaObject subject, uint eventMask, object param)
        {
            if (!componentEnabled || !handleEvents)
            {
                return;
            }
            if (GMEConsole == null)
            {
                GMEConsole = GMEConsole.CreateFromProject(subject.Project);
                //if (consoleReference.GMEConsole == null)
                //{
                //    consoleReference.GMEConsole = GMEConsole;
                //    if (consoleReference.GMEConsole.gme == null)
                //    {
                //        consoleReference.GMEConsole.gme = GMEConsole.gme;
                //    }
                //}
            }
            if ((eventMask & (uint)objectevent_enum.OBJEVENT_OPENMODEL) != 0)
            {
                openModels.Add(subject.AbsPath);
#if(DEBUG)
                GMEConsole.Error.Write(String.Format("Opened Model: {0}", subject.AbsPath));
#endif
            }
            if ((eventMask & (uint)objectevent_enum.OBJEVENT_CLOSEMODEL) != 0)
            {
                openModels.Remove(subject.AbsPath);
#if(DEBUG)
                GMEConsole.Error.Write(String.Format("Closed Model: {0}", subject.AbsPath));
#endif
            }

            // TODO: Handle object events (OR eventMask with the members of objectevent_enum)
            // Warning: Only those events are received that you have subscribed for by setting ComponentConfig.eventMask

            // If the event is OBJEVENT_DESTROYED, most operations on subject will fail
            //   Safe operations: getting Project, ObjType, ID, MetaRole, Meta, MetaBase, Name, AbsPath
            //   Operations that will fail: all others, including attribute access and graph navigation
            //     Try handling OBJEVENT_PRE_DESTROYED if these operations are necessary

            // Be careful not to modify Library objects (check subject.IsLibObject)

            // MessageBox.Show(eventMask.ToString());
            // GMEConsole.Out.WriteLine(subject.Name);

        }

        #endregion

        #region IMgaComponentEx Members

        public void Initialize(MgaProject project)
        {
            // Creating addon
            project.CreateAddOn(this, out addon);
            this.project = project;
            // Setting event mask (see ComponentConfig.eventMask)
            unchecked
            {
                addon.EventMask = (uint)ComponentConfig.eventMask;
            }
            hiddenWindow = new System.Windows.Forms.Control();
            IntPtr handle = hiddenWindow.Handle; // If the handle has not yet been created, referencing this property will force the handle to be created.
            //if (consoleReference == null)
            //{
            //    consoleReference = new GMEConsoleReference();
            //}
        }

        public void InvokeEx(MgaProject project, MgaFCO currentobj, MgaFCOs selectedobjs, int param)
        {
            throw new NotImplementedException(); // Not called by addon
        }


        #region Component Information
        public string ComponentName
        {
            get { return GetType().Name; }
        }

        public string ComponentProgID
        {
            get
            {
                return ComponentConfig.progID;
            }
        }

        public componenttype_enum ComponentType
        {
            get { return ComponentConfig.componentType; }
        }
        public string Paradigm
        {
            get { return ComponentConfig.paradigmName; }
        }
        #endregion

        #region Enabling
        bool enabled = true;
        public void Enable(bool newval)
        {
            enabled = newval;
        }
        #endregion

        #region Interactive Mode
        protected bool interactiveMode = true;
        public bool InteractiveMode
        {
            get
            {
                return interactiveMode;
            }
            set
            {
                interactiveMode = value;
            }
        }
        #endregion

        #region Custom Parameters
        SortedDictionary<string, object> componentParameters = null;

        public object get_ComponentParameter(string Name)
        {
            if (Name == "type")
                return "csharp";

            if (Name == "path")
                return GetType().Assembly.Location;

            if (Name == "fullname")
                return GetType().FullName;

            object value;
            if (componentParameters != null && componentParameters.TryGetValue(Name, out value))
            {
                return value;
            }

            return null;
        }

        public void set_ComponentParameter(string Name, object pVal)
        {
            if (componentParameters == null)
            {
                componentParameters = new SortedDictionary<string, object>();
            }

            componentParameters[Name] = pVal;
        }
        #endregion

        #region Unused Methods
        // Old interface, it is never called for MgaComponentEx interfaces
        public void Invoke(MgaProject Project, MgaFCOs selectedobjs, int param)
        {
            throw new NotImplementedException();
        }

        // Not used by GME
        public void ObjectsInvokeEx(MgaProject Project, MgaObject currentobj, MgaObjects selectedobjs, int param)
        {
            throw new NotImplementedException();
        }

        #endregion

        #endregion

        #region IMgaVersionInfo Members

        public GMEInterfaceVersion_enum version
        {
            get { return GMEInterfaceVersion_enum.GMEInterfaceVersion_Current; }
        }

        #endregion

        #region Registration Helpers

        [ComRegisterFunctionAttribute]
        public static void GMERegister(Type t)
        {
            Registrar.RegisterComponentsInGMERegistry();
        }

        [ComUnregisterFunctionAttribute]
        public static void GMEUnRegister(Type t)
        {
            Registrar.UnregisterComponentsInGMERegistry();
        }

        #endregion

        //public void Dispose()
        //{
        //    if (addon != null)
        //    {
        //        addon.Destroy();
        //        addon = null;
        //    }
        //}
    }
}
