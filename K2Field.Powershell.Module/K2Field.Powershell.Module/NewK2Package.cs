﻿using System;
using System.IO;
using System.Management.Automation;
using SourceCode.ComponentModel;
using SourceCode.Deployment.Management;

namespace K2Field.Powershell.Module
{
    [Cmdlet(VerbsCommon.New, "K2Package")]
    public class NewK2Package : Cmdlet
    {
        private readonly PackageDeploymentManager _packageDeploymentManager = new PackageDeploymentManager();

        [Parameter(Mandatory = true, Position = 1,
            HelpMessage = "Full path to the kspx file, which should appear in the end.")]
        public string Path { get; set; }

        [Parameter(
            Mandatory = true,
            Position = 2,
            HelpMessage = "Workflow(s) e.g. ncl.Rota.Processes\\*"
        )]
        public string Name { get; set; }

        [Parameter(
            Mandatory = true,
            Position = 3,
            HelpMessage = "Namespace e.g. urn:SourceCode/Workflows"
        )]
        public string NameSpace { get; set; }

        [Parameter(Position = 4,
            HelpMessage = "Should the package be validated in the end")]
        public SwitchParameter Validate { get; set; }

        [Parameter(
            Position = 5,
            HelpMessage = "$false: Dependencies will be included into the package, $true: Dependencies will be included as reference.")]
        public SwitchParameter IncludeDependenciesAsReference { get; set; }

        protected override void BeginProcessing()
        {
            try
            {
                _packageDeploymentManager.CreateConnection();
                _packageDeploymentManager.Connection.Open(ModuleHelper.BuildConnectionString());
            }
            catch (Exception ex)
            {
                ErrorHelper.Write(ex);
            }
        }
        protected override void ProcessRecord()
        {
            try
            {
                Session session = _packageDeploymentManager.CreateSession("K2Module" + DateTime.Now.ToString());

                session.SetOption("NoAnalyze", true);
                PackageItemOptions options = PackageItemOptions.Create();
                options.ValidatePackage = Validate;

                var typeRef = new TypeRef(Name, NameSpace);
                
                var query = QueryItemOptions.Create(typeRef);

                var results = session.FindItems(query).Result;
                foreach (var result in results)
                {
                    options.Include(result, PackageItemMode.IncludeDependencies);
                }
                options.PackageItemMode = IncludeDependenciesAsReference ? PackageItemMode.IncludeDependenciesAsReference : PackageItemMode.IncludeDependencies;
                session.PackageItems(options);
                var fileStream = new FileStream(Path, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
                session.Model.Save(fileStream);

                fileStream.Close();
                _packageDeploymentManager.CloseSession(session.Name);
            }
            catch (Exception ex)
            {
                ErrorHelper.Write(ex);
            }
        }
        protected override void EndProcessing()
        {
            _packageDeploymentManager.Connection?.Close();
        }

        protected override void StopProcessing()
        {
            _packageDeploymentManager.Connection?.Close();
        }
    }
}
