using System.Windows.Threading;
using System.IO;
using BckPPT.Services;
using BckPPT.Llm;
using BckPPT.PowerPoint;
using BckPPT.Infrastructure;
using BckPPT.Rag;
using BckPPT.ViewModels;
using BckPPT.Views;

namespace BckPPT.Application;

public sealed class AppCompositionRoot : IDisposable
{
    private readonly IPptComService _pptService = new PptComService();
    private readonly IUiDispatcher _uiDispatcher;
    private readonly DocumentSession _documentSession = new();
    private readonly ILlmClient _llmClient;
    private readonly ApprovalWorkflow _approvalWorkflow;
    private readonly AgentWorkflow _agentWorkflow;
    private readonly TemporaryArtifactStore _temporaryArtifacts = new();
    private readonly PptIndexer _pptIndexer;
    private readonly ExternalJobRegistry _externalJobs = new();
    private ExternalCommandEndpoint? _externalEndpoint;
    private readonly UpdateService _updateService = new();
    private readonly StartupOptions _startupOptions;
    private BckPptWindow? _editorWindow;

    public AppCompositionRoot(Dispatcher dispatcher, StartupOptions? startupOptions = null)
    {
        _startupOptions = startupOptions ?? StartupOptions.Empty;
        _uiDispatcher = new WpfDispatcher(dispatcher);
        GitokyLlmOptions? llmOptions = null;
        try
        {
            llmOptions = GitokyLlmOptions.FromEnvironment(AppContext.BaseDirectory);
            _llmClient = new GitokyLlmClient(llmOptions);
        }
        catch (Exception ex) { _llmClient = new UnavailableLlmClient(ex.Message); }
        var databasePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "BckSoft", "BckPPT", "templates.db");
        var repository = new TemplateRepository(databasePath);
        var retriever = new TemplateRetriever(repository);
        _pptIndexer = new PptIndexer(repository);
        var renderValidator = new RenderValidator(_pptService, _temporaryArtifacts);
        _approvalWorkflow = new ApprovalWorkflow(_pptService, renderValidator, new LayoutDecisionRouter(_pptService, retriever));
        _agentWorkflow = new AgentWorkflow(_llmClient, llmOptions, _pptService, renderValidator,
            new SourceLibrary(SourceLibrary.DefaultDatabasePath));
    }

    public AgentChatWindow CreateAgentChatWindow() =>
        new(_documentSession, _pptService, _llmClient, _approvalWorkflow, _agentWorkflow, _pptIndexer,
            _externalJobs, _updateService, OpenEditorWindow, _startupOptions.ResolveInitialDirectory());

    public void StartExternalEndpoint(AgentChatWindow window)
    {
        _externalEndpoint = new ExternalCommandEndpoint(_externalJobs, request =>
        {
            window.Dispatcher.BeginInvoke(() => window.AcceptExternalRequest(request));
            return Task.CompletedTask;
        });
        _externalEndpoint.Start();
    }

    private void OpenEditorWindow()
    {
        if (_editorWindow is { IsLoaded: true })
        {
            _editorWindow.Activate();
            return;
        }

        var viewModel = new BckPptViewModel(_pptService, _uiDispatcher);
        _editorWindow = new BckPptWindow(viewModel, _documentSession, _startupOptions.ResolveInitialDirectory());
        _editorWindow.Closed += (_, _) => _editorWindow = null;
        _editorWindow.Show();
    }

    public void Dispose()
    {
        _agentWorkflow.Dispose();
        (_llmClient as IDisposable)?.Dispose();
        _externalEndpoint?.Dispose();
        _updateService.Dispose();
        _temporaryArtifacts.Dispose();
    }
}

public sealed class WpfDispatcher(Dispatcher dispatcher) : IUiDispatcher
{
    public void Post(Action action) => dispatcher.BeginInvoke(action);
}
