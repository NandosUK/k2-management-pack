using System;
using System.Globalization;
using System.Management.Automation;
using SourceCode.EnvironmentSettings.Client;
using SourceCode.Framework.Deployment;
using SourceCode.Hosting.Client.BaseAPI;
using SourceCode.ProjectSystem;

namespace K2Field.Powershell.Module
{
    [Cmdlet(VerbsCommon.New, "K2PackageBuild")]
    public class NewK2PackageBuild : Cmdlet
    {
        [Parameter(Mandatory = true, Position = 1, HelpMessage = "Full path to the K2.proj file")]
        public string ProjectPath { get; set; }


        [Parameter(Mandatory = true, Position = 2, HelpMessage = "The K2 Environment Template ")]
        public string Template { get; set; }

        [Parameter(Mandatory = true, Position = 3, HelpMessage = "The default deployment environment")]
        public string Environment { get; set; }

        [Parameter(Mandatory = true, Position = 4, HelpMessage = "The target server")]
        public string Server { get; set; }


        protected override void ProcessRecord()
        {
            var project = new Project();
            project.Load(this.ProjectPath);

            var result2 = project.Compile();

            if (!result2.Successful)
            {
                var message = result2.Errors[0].ErrorText;

                throw new Exception(message);
            }

            var package = GetPackage(this.Server, this.Template, this.Environment, project, false);

            var result = package.Execute();

            if (!result.Successful)
            {
                throw new Exception("Error occurred deploying package");
            }
        }

        private static string EnvionmentServerConnection(string environmentServer)
        {
            SCConnectionStringBuilder cb = new SCConnectionStringBuilder();
            cb.Host = environmentServer;
            cb.Port = 5555;
            cb.Integrated = true;
            cb.IsPrimaryLogin = true;
            cb.Authenticate = true;
            cb.EncryptedPassword = false;
            string envionmentServerConnection = cb.ToString();
            return envionmentServerConnection;
        }

        public static DeploymentPackage GetPackage(string environmentServer,
           string destinationTemplate, string destinationEnvironment,
           Project project, bool testOnly)
        {
            //Create connection string to environment server
            var envionmentServerConnection = EnvionmentServerConnection(environmentServer);
            //Retrieve the environments from the server
            EnvironmentSettingsManager environmentManager = new
                EnvironmentSettingsManager(true);
            environmentManager.ConnectToServer(envionmentServerConnection);
            environmentManager.InitializeSettingsManager(true);
            environmentManager.Refresh();

            //Get the template and environment objects.
            EnvironmentTemplate template = environmentManager.EnvironmentTemplates[destinationTemplate];
            EnvironmentInstance environment = template.Environments[destinationEnvironment];
            //Create the package
            DeploymentPackage package = project.CreateDeploymentPackage();

            //Set all of the environment fields to the package
            DeploymentEnvironment deploymentEnv =
                package.AddEnvironment(environment.EnvironmentName);
            foreach (EnvironmentField field in environment.EnvironmentFields)
            {
                deploymentEnv.Properties[field.FieldName] = field.Value;
            }

            //Set fields on the package
            package.SelectedEnvironment = destinationEnvironment;
            package.DeploymentLabelName = DateTime.Now.ToString(CultureInfo.InvariantCulture);
            package.DeploymentLabelDescription =
                "Template: " + destinationTemplate + ",Environment: " + destinationEnvironment;
            package.TestOnly = testOnly;
            //Get the Default SmartObject Server in the Environment
            //environment.GetDefaultField(typeof(SmartObjectField));
            package.SmartObjectConnectionString = envionmentServerConnection;
            //Get the Default Workflow Management Server in the Environment
            //environment.GetDefaultField(typeof(WorkflowManagementServerField));
            package.WorkflowManagementConnectionString = envionmentServerConnection;

            return package;
        }
    }
}