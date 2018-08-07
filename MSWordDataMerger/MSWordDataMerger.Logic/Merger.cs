using System;
using System.IO;

namespace MSWordDataMerger.Logic
{
    public class Merger
    {
        private IKeyValueLoader keyValueLoader = null;
        private ITemplateEditor templateEditor = null;
        private String templateRootDirectory;

        public Merger(KeyValuePairLoaderType keyValuePairLoaderType, TemplateEditorType templateEditorType, String templateRootDirectory)
        {
            switch (keyValuePairLoaderType)
            {
                case KeyValuePairLoaderType.SeperateFileBased:
                    keyValueLoader = new SeperateFileBasedKeyValueLoader();
                    break;
                default:
                    throw new ArgumentException("Unknown KeyValuePairLoaderType");
            }

            switch (templateEditorType)
            {
                case TemplateEditorType.XceedSoftwareDocX:
                    templateEditor = new DocXTemplateEditor();
                    break;
                default:
                    throw new ArgumentException("Unknown TemplateEditorType");
            }

            this.templateRootDirectory = templateRootDirectory;
        }

        public void Merge()
        {
            if (Directory.Exists(templateRootDirectory))
            {
                var templatePath = GetTemplatePath(templateRootDirectory);
                var dataSource = $@"{templateRootDirectory}\mergeRequests";
                if (templatePath != "" && Directory.Exists(dataSource) &&
                    Directory.GetDirectories(dataSource).Length > 0)
                {
                    var dataDirectories = Directory.GetDirectories(dataSource);
                    foreach (var dataDirectory in dataDirectories)
                    {
                        if (File.Exists($@"{dataDirectory}\finishedExport"))
                        {
                            try
                            {
                                // Merge template with data
                                var data = keyValueLoader.LoadKeyValuePairs($@"{dataDirectory}\data");

                                templateEditor.MergeWithData(
                                    templatePath: templatePath,
                                    keyValuePairs: data.KeyValuePairs,
                                    iterationKeyValuePairHolders: data.KeyValuePairIterations,
                                    toAppendDocumentsPath: AppendDocuments(dataDirectory)
                                        ? $@"{templateRootDirectory}\appendDocuments"
                                        : "",
                                    mergedDocOutputPath:
                                    $@"{templateRootDirectory}\output\{GetOuputName(dataDirectory)}"
                                );
                            }
                            catch (MergerException)
                            {
                                throw;  // Rethrow this exception so it can be handled by the class that's using this one
                            }
                            catch (Exception e)
                            {
                                throw new MergerException("Something went wrong while trying to merge data with the template.", e.Message);
                                // TODO Possible memoryleak? Will it reach the finally block?
                                // Also possible loop if it does not reach the finally block
                            }
                            finally
                            {           // TODO I think it does get here....
                                try
                                {
                                    Directory.Delete(dataDirectory, true);
                                }
                                catch (Exception e)
                                {
                                    throw new MergerException("Could not delete the data directory from the mergeRequests folder.", e.Message);
                                }
                            }


                        }
                    }
                }
            }
            else
            {
                throw new MergerException("The given template root directory does not exist.", "Does the service have access to that folder?");
            }
        }

        private bool AppendDocuments(String dataIterationPath)
        {
            return File.Exists($@"{dataIterationPath}\config\appendDocuments");
        }

        private String GetOuputName(String dataIterationPath)
        {
            var name = $"{new DirectoryInfo(dataIterationPath).Name}.docx";
            var settingFilePath = $@"{dataIterationPath}\config\outputName";
            try
            {
                var readData = File.ReadAllText(settingFilePath);
                name = readData;
            }
            catch (Exception e)
            {
                throw new MergerException(
                    $"Could not read the content of the output setting file located at \"{settingFilePath}\"",
                    e.Message);
            }
            return name;
        }

        private String GetTemplatePath(String templateDirectory)
        {
            var path = "";
            var files = Directory.GetFiles(templateDirectory);
            foreach (var file in files)
            {
                if (Path.GetExtension(file) == ".docx")
                {
                    return file;
                }
            }
            return path;
        }
    }
}
