using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace VSPremake
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class PremakeCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("efa1febb-39cc-4455-8d50-1b50fe3c5e5c");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Location of premake.exe
        /// </summary>
        private string premakeExecutableLocation = string.Empty;

        /// <summary>
        /// Initializes a new instance of the <see cref="PremakeCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private PremakeCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);

            premakeExecutableLocation = Path.GetTempFileName();
            File.WriteAllBytes(premakeExecutableLocation, Properties.Resources.premake5);

        }

        ~PremakeCommand()
        {
            File.Delete(premakeExecutableLocation);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static PremakeCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in PremakeCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new PremakeCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
#pragma warning disable VSTHRD100 // Avoid async void methods
        private async void Execute(object sender, EventArgs e)
#pragma warning restore VSTHRD100 // Avoid async void methods
        {
            try
            {
                //Ensure that we are on main thread
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

                var outputWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;

                if (outputWindow != null)
                {
                    var guidGeneral = new Guid("519B77C0-1DCC-4561-AB6C-3181A1B75A6C");
                    IVsOutputWindowPane pane;
                    outputWindow.CreatePane(guidGeneral, "Premake", 1, 0);
                    outputWindow.GetPane(guidGeneral, out pane);
                    pane.Activate();
                    pane.Clear();

                    //Need this to get user arguments
                    VSPremakePackage package = this.package as VSPremakePackage;

                    //Need this to get the path of the premake file
                    IVsSolution solution = (IVsSolution)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(IVsSolution));
                    solution.GetSolutionInfo(out string solutionDirectory, out string solutionName, out string solutionDirectory2);

                    var proc = new System.Diagnostics.Process();
                    proc.StartInfo.CreateNoWindow = true;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.StartInfo.RedirectStandardError = true;
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.FileName = premakeExecutableLocation;
                    proc.StartInfo.WorkingDirectory = solutionDirectory;
                    proc.StartInfo.Arguments = package.OptionArguments;
                    proc.Start();

                    proc.OutputDataReceived += (o, args) =>
                    {
#pragma warning disable VSTHRD010 // Invoke single-threaded types on Main thread
                        _ = pane.OutputStringThreadSafe(args.Data + "\n");
#pragma warning restore VSTHRD010 // Invoke single-threaded types on Main thread
                    };
                    proc.ErrorDataReceived += (o, args) =>
                    {
#pragma warning disable VSTHRD010 // Invoke single-threaded types on Main thread
                        _ = pane.OutputStringThreadSafe(args.Data + "\n");
#pragma warning restore VSTHRD010 // Invoke single-threaded types on Main thread
                    };

                    proc.BeginOutputReadLine();
                    proc.BeginErrorReadLine();

                }
            }
            catch (Exception)
            {
                ; //TODO: Should probably alert the user if something fails, but I cant be bothered.
            }
        }
    }
}
