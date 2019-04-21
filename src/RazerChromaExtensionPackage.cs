using System;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Task = System.Threading.Tasks.Task;

using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

using EnvDTE;
using EnvDTE80;
 
using Corale.Colore.Core;

namespace RazerChromaExtension
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>    
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(RazerChromaExtensionPackage.PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class RazerChromaExtensionPackage : AsyncPackage
    {
        private bool _isBuildSucceeded = false;
        private DTE2 _dte;

        /// <summary>
        /// Razer_Chroma_ExtensionPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "1cd39cbb-d621-439b-93a0-2a304e7136c2";

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {

            // Since this package might not be initialized until after a solution has finished loading,
            // we need to check if a solution has already been loaded and then handle it.
            bool isSolutionLoaded = await IsSolutionLoadedAsync(cancellationToken);

            if (isSolutionLoaded)
            {
                HandleOpenSolution();
            }

            // Listen for subsequent solution events
            Microsoft.VisualStudio.Shell.Events.SolutionEvents.OnAfterAsynchOpenProject += HandleOpenSolution;
        }
        
        private async void HandleOpenSolution(object sender = null, EventArgs e = null)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync();

            _dte.Events.BuildEvents.OnBuildBegin += BuildEvents_OnBuildBegin;
            _dte.Events.BuildEvents.OnBuildProjConfigDone += BuildEvents_OnBuildProjConfigDone;
            _dte.Events.BuildEvents.OnBuildProjConfigBegin += BuildEvents_OnBuildProjConfigBegin;
            _dte.Events.BuildEvents.OnBuildDone += BuildEvents_OnBuildDone;
        }

        private async Task<bool> IsSolutionLoadedAsync(CancellationToken cancellationToken)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            this._dte = await GetServiceAsync(typeof(DTE)).ConfigureAwait(false) as DTE2;
            Assumes.Present(this._dte);

            var solService = await GetServiceAsync(typeof(SVsSolution)) as IVsSolution;
            Assumes.Present(solService);

            ErrorHandler.ThrowOnFailure(solService.GetProperty((int)__VSPROPID.VSPROPID_IsSolutionOpen, out object value));

            return value is bool isSolOpen && isSolOpen;
        }

        private void BuildEvents_OnBuildProjConfigBegin(string project, string projectConfig, string platform, string solutionConfig)
        {
            Chroma.Instance.SetAll(Color.White);
            System.Threading.Thread.Sleep(850);
        }

        private void BuildEvents_OnBuildProjConfigDone(string project, string projectConfig, string platform, string solutionConfig, bool success)
        {
            this._isBuildSucceeded = success;
        }

        private void BuildEvents_OnBuildDone(vsBuildScope scope, vsBuildAction action)
        {
            Color lightingColor = _isBuildSucceeded ? Color.Green : Color.Red;

            for (int index = 0; index < 10; ++index)
            {
                System.Threading.Thread.Sleep(80);
                Chroma.Instance.SetAll(lightingColor);
                System.Threading.Thread.Sleep(250);
                Chroma.Instance.SetAll(Color.Black);
            }

            System.Threading.Thread.Sleep(500);
            Chroma.Instance.Uninitialize();

        }

        private void BuildEvents_OnBuildBegin(vsBuildScope scope, vsBuildAction action)
        {
            Chroma.Instance.Initialize();
            Chroma.Instance.SetAll(Color.White);

            System.Threading.Thread.Sleep(500);
        }
    }
    #endregion
}
