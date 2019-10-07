/////////////////////////////////////////////////////////////////////
// Copyright (c) Autodesk, Inc. All rights reserved
// Written by Forge Partner Development
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
/////////////////////////////////////////////////////////////////////

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

using Inventor;
using Autodesk.Forge.DesignAutomation.Inventor.Utils;
using Newtonsoft.Json;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

using System.IO.Compression;
using File = System.IO.File;
using Path = System.IO.Path;
using Directory = System.IO.Directory;

namespace ExtractParamsPlugin
{
    [ComVisible(true)]
    public class SampleAutomation
    {
        private readonly InventorServer inventorApplication;
        public bool IsLocalDebug = false;

        public SampleAutomation(InventorServer inventorApp)
        {
            inventorApplication = inventorApp;
        }

        public void Run(Document doc)
        {
            LogTrace("Run()");

            if (IsLocalDebug)
            {
                dynamic dynDoc = doc;
                string parameters = getParamsAsJson(dynDoc.ComponentDefinition.Parameters.UserParameters);

                System.IO.File.WriteAllText("documentParams.json", parameters);

                // Save Forge Viewable
                CreateViewable(doc);
            }
            else
            {
                string currentDir = System.IO.Directory.GetCurrentDirectory();
                LogTrace("Current Dir = " + currentDir);

                // Helpful for debugging input files
                //LogDir(currentDir);

                Dictionary<string, string> inputParameters = JsonConvert.DeserializeObject<Dictionary<string, string>>(System.IO.File.ReadAllText("inputParams.json"));
                logInputParameters(inputParameters);

                using (new HeartBeat())
                {
                    if (inputParameters.ContainsKey("projectPath"))
                    {
                        string projectPath = inputParameters["projectPath"];
                        string fullProjectPath = System.IO.Path.GetFullPath(System.IO.Path.Combine(currentDir, projectPath));
                        LogTrace("fullProjectPath = " + fullProjectPath);
                        DesignProject dp = inventorApplication.DesignProjectManager.DesignProjects.AddExisting(fullProjectPath);
                        dp.Activate();
                    }
                    else
                    {
                        LogTrace("No 'projectPath' property");
                    }

                    if (inputParameters.ContainsKey("documentPath"))
                    {
                        string documentPath = inputParameters["documentPath"];

                        string fullDocumentPath = System.IO.Path.GetFullPath(System.IO.Path.Combine(currentDir, documentPath));
                        LogTrace("fullDocumentPath = " + fullDocumentPath);
                        dynamic invDoc = inventorApplication.Documents.Open(fullDocumentPath);
                        LogTrace("Opened input document file");
                        dynamic compDef = invDoc.ComponentDefinition;

                        string parameters = getParamsAsJson(compDef.Parameters.UserParameters);

                        System.IO.File.WriteAllText("documentParams.json", parameters);
                        LogTrace("Created documentParams.json");

                        // Save out Forge Viewable
                        CreateViewable(invDoc);
                    }
                    else
                    {
                        LogTrace("No 'documentPath' property");
                    }
                }
            }
        }

        static void LogDir(string dir)
        {
            try
            {
                foreach (string d in Directory.GetDirectories(dir))
                {
                    foreach (string f in Directory.GetFiles(d))
                    {
                        LogTrace(f);
                    }
                    LogDir(d);
                }
            }
            catch (System.Exception excpt)
            {
                LogTrace(excpt.Message);
            }
        }

        public string getParamsAsJson(dynamic userParameters)
        {
            /* The resulting json will be like this:
              { 
                "length" : {
                  "unit": "in",
                  "value": "10 in",
                  "values": ["5 in", "10 in", "15 in"]
                },
                "width": {
                  "unit": "in",
                  "value": "20 in",
                }
              }
            */
            List<object> parameters = new List<object>();
            foreach (dynamic param in userParameters)
            {
                List<object> paramProperties = new List<object>();
                if (param.ExpressionList != null)
                {
                    string[] expressions = param.ExpressionList.GetExpressionList();
                    JArray values = new JArray(expressions);
                    paramProperties.Add(new JProperty("values", values));
                }
                paramProperties.Add(new JProperty("value", param.Expression));
                paramProperties.Add(new JProperty("unit", param.Units));

                parameters.Add(new JProperty(param.Name, new JObject(paramProperties.ToArray())));
            }
            JObject allParameters = new JObject(parameters.ToArray());
            string paramsJson = allParameters.ToString();
            LogTrace(paramsJson);

            return paramsJson;
        }

        public void logInputParameters(Dictionary<string, string> parameters)
        {
            foreach (KeyValuePair<string, string> entry in parameters)
            {
                try
                {
                    LogTrace("Key = {0}, Value = {1}", entry.Key, entry.Value);
                }
                catch (Exception e)
                {
                    LogTrace("Error with key {0}: {1}", entry.Key, e.Message);
                }
            }
        }

        private void CreateViewable(Document doc)
        {
            // Save out Forge Viewable
            var docDir = Path.GetDirectoryName(doc.FullFileName);
            string viewableDir = SaveForgeViewable(doc);
            ZipOutput(viewableDir, "viewable.zip");
        }

        private string SaveForgeViewable(Document doc)
        {
            string viewableOutputDir = null;
            using (new HeartBeat())
            {
                LogTrace($"** Saving SVF");
                try
                {
                    TranslatorAddIn oAddin = null;


                    foreach (ApplicationAddIn item in inventorApplication.ApplicationAddIns)
                    {

                        if (item.ClassIdString == "{C200B99B-B7DD-4114-A5E9-6557AB5ED8EC}")
                        {
                            Trace.TraceInformation("SVF Translator addin is available");
                            oAddin = (TranslatorAddIn)item;
                            break;
                        }
                        else { }
                    }

                    if (oAddin != null)
                    {
                        Trace.TraceInformation("SVF Translator addin is available");
                        TranslationContext oContext = inventorApplication.TransientObjects.CreateTranslationContext();
                        // Setting context type
                        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;

                        NameValueMap oOptions = inventorApplication.TransientObjects.CreateNameValueMap();
                        // Create data medium;
                        DataMedium oData = inventorApplication.TransientObjects.CreateDataMedium();

                        Trace.TraceInformation("SVF save");
                        var workingDir = Path.GetDirectoryName(doc.FullFileName);
                        var sessionDir = Path.Combine(workingDir, "SvfOutput");

                        // Make sure we delete any old contents that may be in the output directory first,
                        // this is for local debugging. In DA4I the working directory is always clean
                        if (Directory.Exists(sessionDir))
                        {
                            Directory.Delete(sessionDir, true);
                        }

                        oData.FileName = Path.Combine(sessionDir, "result.collaboration");
                        var outputDir = Path.Combine(sessionDir, "output");
                        var bubbleFileOriginal = Path.Combine(outputDir, "bubble.json");
                        var bubbleFileNew = Path.Combine(sessionDir, "bubble.json");

                        // Setup SVF options
                        if (oAddin.get_HasSaveCopyAsOptions(doc, oContext, oOptions))
                        {
                            oOptions.set_Value("GeometryType", 1);
                            oOptions.set_Value("EnableExpressTranslation", true);
                            oOptions.set_Value("SVFFileOutputDir", sessionDir);
                            oOptions.set_Value("ExportFileProperties", false);
                            oOptions.set_Value("ObfuscateLabels", true);
                        }

                        LogTrace($"SVF files are oputput to: {oOptions.get_Value("SVFFileOutputDir")}");

                        oAddin.SaveCopyAs(doc, oContext, oOptions, oData);
                        Trace.TraceInformation("SVF can be exported.");
                        LogTrace($"** Saved SVF as {oData.FileName}");
                        File.Move(bubbleFileOriginal, bubbleFileNew);

                        viewableOutputDir = sessionDir;
                    }
                }
                catch (Exception e)
                {
                    LogError($"********Export to format SVF failed: {e.Message}");
                    return null;
                }
            }
            return viewableOutputDir;
        }

        private void ZipOutput(string pathName, string fileName)
        {
            try
            {
                LogTrace($"Zipping up {fileName}");

                if (File.Exists(fileName)) File.Delete(fileName);

                // start HeartBeat around ZipFile, it could be a long operation
                using (new HeartBeat())
                {
                    ZipFile.CreateFromDirectory(pathName, fileName, CompressionLevel.Fastest, false);
                }

                LogTrace($"Saved as {fileName}");
            }
            catch (Exception e)
            {
                LogError($"********Export to format SVF failed: {e.Message}");
            }
        }

        #region Logging utilities

        /// <summary>
        /// Log message with 'trace' log level.
        /// </summary>
        private static void LogTrace(string format, params object[] args)
        {
            Trace.TraceInformation(format, args);
        }

        /// <summary>
        /// Log message with 'trace' log level.
        /// </summary>
        private static void LogTrace(string message)
        {
            Trace.TraceInformation(message);
        }

        /// <summary>
        /// Log message with 'error' log level.
        /// </summary>
        private static void LogError(string format, params object[] args)
        {
            Trace.TraceError(format, args);
        }

        /// <summary>
        /// Log message with 'error' log level.
        /// </summary>
        private static void LogError(string message)
        {
            Trace.TraceError(message);
        }

        #endregion
    }
}