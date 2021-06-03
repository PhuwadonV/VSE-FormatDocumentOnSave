using System;
using System.IO;
using System.Text.RegularExpressions;
using Elders.VSE_FormatDocumentOnSave.Configurations;
using EnvDTE;

namespace Elders.VSE_FormatDocumentOnSave
{
    public class DocumentFormatService
    {
        private static readonly Regex MACRO_REGEX = new Regex(@"^(?<leadingWhitespace>[\s]*)(UPROPERTY|UFUNCTION)\(.*");
        private readonly DTE dte;
        readonly Func<Document, IConfiguration> getGeneralCfg;
        readonly IDocumentFormatter formatter;

        public DocumentFormatService(DTE dte, Func<Document, IConfiguration> getGeneralCfg)
        {
            this.dte = dte;
            this.getGeneralCfg = getGeneralCfg;

            formatter = new VisualStudioCommandFormatter(dte);
        }

        public void FormatDocument(Document doc)
        {
            if (ShouldFormat(doc) == false)
                return;

            try
            {
                var cfg = getGeneralCfg(doc);
                var filter = new AllowDenyDocumentFilter(cfg.Allowed, cfg.Denied);

                foreach (string splitCommand in cfg.Commands.Trim().Split(' '))
                {
                    try
                    {
                        string commandName = splitCommand.Trim();
                        formatter.Format(doc, filter, commandName);
                    }
                    catch (Exception) { }   // may be we can log which command has failed and why
                }

                FileInfo fileInfo = new FileInfo(doc.FullName);
                if (fileInfo.Extension == ".h")
                {
                    var dte = doc.DTE;
                    dte.UndoContext.Open("FixUnrealEngineMacroIndents");

                    try
                    {
                        var textDoc = (TextDocument)dte.ActiveDocument.Object("TextDocument");
                        var underMacro = false;
                        string leadingWhitespace = null;

                        for (var editPoint = textDoc.StartPoint.CreateEditPoint(); !editPoint.AtEndOfDocument; editPoint.LineDown())
                        {
                            if (!underMacro)
                            {
                                var lineNum = editPoint.Line;
                                var line = editPoint.GetLines(lineNum, lineNum + 1);
                                var macroMatch = MACRO_REGEX.Match(line);
                                underMacro = macroMatch.Success;
                                leadingWhitespace = macroMatch.Groups["leadingWhitespace"].ToString();
                            }
                            else
                            {
                                editPoint.StartOfLine();
                                editPoint.DeleteWhitespace(EnvDTE.vsWhitespaceOptions.vsWhitespaceOptionsHorizontal);
                                editPoint.Insert(leadingWhitespace);
                                underMacro = false;
                            }
                        }
                    }
                    finally
                    {
                        dte.UndoContext.Close();
                    }
                }
            }
            catch (Exception) { }   // Do not do anything here on purpose.
        }

        private bool ShouldFormat(Document doc)
        {
            if (System.Windows.Forms.Control.IsKeyLocked(System.Windows.Forms.Keys.CapsLock))
                return false;

            bool vsIsInDebug = dte.Mode == vsIDEMode.vsIDEModeDebug;
            var cfg = getGeneralCfg(doc);

            if (vsIsInDebug == true && cfg.EnableInDebug == false)
                return false;

            return cfg.IsEnable;
        }
    }
}