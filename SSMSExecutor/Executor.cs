using System;
using System.Collections.Generic;

using EnvDTE;
using EnvDTE80;
using Microsoft.SqlServer.TransactSql.ScriptDom;

namespace Devvcat.SSMS
{
    sealed class Executor
    {
        public readonly string CMD_QUERY_EXECUTE = "Query.Execute";

        private Document document;

        private EditPoint oldAnchor;
        private EditPoint oldActivePoint;

        public Executor(DTE2 dte)
        {
            if (dte == null) throw new ArgumentNullException(nameof(dte));

            document = dte.GetDocument();

            SaveActiveAndAnchorPoints();
        }

        private VirtualPoint GetCaretPoint()
        {
            var p = ((TextSelection)document.Selection).ActivePoint;

            return new VirtualPoint(p);
        }

        private string GetDocumentText()
        {
            var content = string.Empty;
            var selection = (TextSelection)document.Selection;

            if (!selection.IsEmpty)
            {
                content = selection.Text;
            }
            else
            {
                if (document.Object("TextDocument") is TextDocument doc)
                {
                    content = doc.StartPoint.CreateEditPoint().GetText(doc.EndPoint);
                }
            }

            return content;
        }

        private void SaveActiveAndAnchorPoints()
        {
            var selection = (TextSelection)document.Selection;

            oldAnchor = selection.AnchorPoint.CreateEditPoint();
            oldActivePoint = selection.ActivePoint.CreateEditPoint();
        }

        private void RestoreActiveAndAnchorPoints()
        {
            var startPoint = new VirtualPoint(oldAnchor);
            var endPoint = new VirtualPoint(oldActivePoint);

            MakeSelection(startPoint, endPoint);
        }

        private void MakeSelection(VirtualPoint startPoint, VirtualPoint endPoint)
        {
            var selection = (TextSelection)document.Selection;

            selection.MoveToLineAndOffset(startPoint.Line, startPoint.LineCharOffset);
            selection.SwapAnchor();
            selection.MoveToLineAndOffset(endPoint.Line, endPoint.LineCharOffset, true);
        }

        private bool ParseSqlFragments(string script, out TSqlScript sqlFragments)
        {
            IList<ParseError> errors;
            TSql140Parser parser = new TSql140Parser(true);

            using (System.IO.StringReader reader = new System.IO.StringReader(script))
            {
                sqlFragments = parser.Parse(reader, out errors) as TSqlScript;
            }

            return errors.Count == 0;
        }

        private IList<TSqlStatement> GetInnerStatements(TSqlStatement statement)
        {
            List<TSqlStatement> list = new List<TSqlStatement>();

            if (statement is BeginEndBlockStatement block)
            {
                list.AddRange(block.StatementList.Statements);
            }
            else if (statement is IfStatement ifBlock)
            {
                if (ifBlock.ThenStatement != null)
                {
                    list.Add(ifBlock.ThenStatement);
                }
                if (ifBlock.ElseStatement != null)
                {
                    list.Add(ifBlock.ElseStatement);
                }
            }
            else if (statement is WhileStatement whileBlock)
            {
                list.Add(whileBlock.Statement);
            }

            return list;
        }

        private bool IsCaretInsideStatement(TSqlStatement statement, VirtualPoint caret)
        {
            var ft = statement.ScriptTokenStream[statement.FirstTokenIndex];
            var lt = statement.ScriptTokenStream[statement.LastTokenIndex];

            if (caret.Line >= ft.Line && caret.Line <= lt.Line)
            {
                var isBeforeFirstToken = caret.Line == ft.Line && caret.LineCharOffset < ft.Column;
                var isAfterLastToken = caret.Line == lt.Line && caret.LineCharOffset > lt.Column + lt.Text.Length;

                if (!(isBeforeFirstToken || isAfterLastToken))
                {
                    return true;
                }
            }

            return false;
        }

        private TextBlock GetTextBlockFromStatement(TSqlStatement statement)
        {
            var ft = statement.ScriptTokenStream[statement.FirstTokenIndex];
            var lt = statement.ScriptTokenStream[statement.LastTokenIndex];

            return new TextBlock()
            {
                StartPoint = new VirtualPoint
                {
                    Line = ft.Line,
                    LineCharOffset = ft.Column
                },

                EndPoint = new VirtualPoint
                {
                    Line = lt.Line,
                    LineCharOffset = lt.Column + lt.Text.Length
                }
            };
        }

        private TextBlock FindCurrentStatement(IList<TSqlStatement> statements, VirtualPoint caret, ExecScope scope)
        {
            if (statements == null || statements.Count == 0)
            {
                return null;
            }

            foreach (var statement in statements)
            {
                if (scope == ExecScope.Inner)
                {
                    IList<TSqlStatement> statementList = GetInnerStatements(statement);

                    TextBlock currentStatement = FindCurrentStatement(statementList, caret, scope);

                    if (currentStatement != null)
                    {
                        return currentStatement;
                    }
                }

                if (IsCaretInsideStatement(statement, caret))
                {
                    return GetTextBlockFromStatement(statement);
                }
            }

            return null;
        }

        private void Exec()
        {
            document.DTE.ExecuteCommand(CMD_QUERY_EXECUTE);
        }

        private bool CanExecute()
        {
            try
            {
                var cmd = document.DTE.Commands.Item(CMD_QUERY_EXECUTE, -1);
                return cmd.IsAvailable;
            }
            catch
            { }

            return false;
        }

        public void ExecuteStatement(ExecScope scope = ExecScope.Block)
        {
            if (!CanExecute())
            {
                return;
            }

            SaveActiveAndAnchorPoints();

            if (!(document.Selection as TextSelection).IsEmpty)
            {
                Exec();
            }
            else
            {
                var script = GetDocumentText();
                var caretPoint = GetCaretPoint();

                bool success = ParseSqlFragments(script, out TSqlScript sqlScript);

                if (success)
                {
                    TextBlock currentStatement = null;

                    foreach (var batch in sqlScript?.Batches)
                    {
                        currentStatement = FindCurrentStatement(batch.Statements, caretPoint, scope);

                        if (currentStatement != null)
                        {
                            break;
                        }
                    }

                    if (currentStatement != null)
                    {
                        // select the statement to be executed
                        MakeSelection(currentStatement.StartPoint, currentStatement.EndPoint);

                        // execute the statement
                        Exec();

                        // restore selection
                        RestoreActiveAndAnchorPoints();
                    }
                }
                else
                {
                    // there are syntax errors
                    // execute anyway to show the errors
                    Exec();
                }
            }
        }

        public class VirtualPoint
        {
            public int Line { get; set; }
            public int LineCharOffset { get; set; }

            public VirtualPoint()
            {
                Line = 1;
                LineCharOffset = 0;
            }

            public VirtualPoint(EnvDTE.TextPoint point)
            {
                Line = point.Line;
                LineCharOffset = point.LineCharOffset;
            }
        }

        public class TextBlock
        {
            public VirtualPoint StartPoint { get; set; }
            public VirtualPoint EndPoint { get; set; }
        }

        internal enum ExecScope
        {
            Block,
            Inner
        }
    }
}
