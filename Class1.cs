#nullable disable

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;

[assembly: CommandClass(typeof(EntityCounter.CountCommand))]

namespace EntityCounter
{
    public class CountCommand
    {
        [CommandMethod("CountEntitiesMulti")]
        public void CountEntitiesMulti()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ed.WriteMessage("\n--- ENTITY COUNTER ---");

            //  STEP 1: Select reference entity
            PromptEntityOptions refOpt = new PromptEntityOptions("\nSelect reference entity: ");
            PromptEntityResult refRes = ed.GetEntity(refOpt);
            if (refRes.Status != PromptStatus.OK) return;

            //  STEP 2: Select objects
            PromptSelectionResult selRes = ed.GetSelection();
            if (selRes.Status != PromptStatus.OK)
            {
                ed.WriteMessage("\nNo objects selected.");
                return;
            }

            //  STEP 3: User input (row height)
            PromptDoubleOptions rowOpt = new PromptDoubleOptions("\nEnter row height: ");
            rowOpt.AllowNegative = false;
            rowOpt.AllowZero = false;
            rowOpt.DefaultValue = 10;

            PromptDoubleResult rowRes = ed.GetDouble(rowOpt);
            if (rowRes.Status != PromptStatus.OK) return;

            //  STEP 3: User input (column width)
            PromptDoubleOptions colOpt = new PromptDoubleOptions("\nEnter column width: ");
            colOpt.AllowNegative = false;
            colOpt.AllowZero = false;
            colOpt.DefaultValue = 50;

            PromptDoubleResult colRes = ed.GetDouble(colOpt);
            if (colRes.Status != PromptStatus.OK) return;

            //  STEP 4: Insertion point
            PromptPointResult ptRes = ed.GetPoint("\nSelect insertion point: ");
            if (ptRes.Status != PromptStatus.OK) return;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    Dictionary<string, int> counts = new Dictionary<string, int>();

                    foreach (SelectedObject so in selRes.Value)
                    {
                        if (so == null) continue;

                        Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;

                        string key = GetKey(ent, tr);

                        if (counts.ContainsKey(key))
                            counts[key]++;
                        else
                            counts.Add(key, 1);
                    }

                    //  CREATE TABLE
                    CreateTable(db, tr, counts, ptRes.Value, rowRes.Value, colRes.Value);

                    ed.WriteMessage("\nTable created successfully.");
                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage("\nError: " + ex.Message);
                    tr.Abort();
                }
            }
        }

        //  Get key
        private string GetKey(Entity ent, Transaction tr)
        {
            if (ent is BlockReference br)
                return "BLOCK_" + GetBlockName(br, tr).ToUpper();

            return ent.GetType().Name.ToUpper();
        }

        //  Get block name
        private string GetBlockName(BlockReference br, Transaction tr)
        {
            try
            {
                if (br.IsDynamicBlock)
                {
                    BlockTableRecord btr =
                        (BlockTableRecord)tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead);
                    return btr.Name;
                }
                else
                {
                    return br.Name;
                }
            }
            catch
            {
                return "UNKNOWN_BLOCK";
            }
        }

        //  CREATE TABLE WITH USER SIZE
        private void CreateTable(Database db, Transaction tr,
            Dictionary<string, int> data, Point3d pt,
            double rowHeight, double colWidth)
        {
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord ms =
                (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            Table tb = new Table();

            tb.TableStyle = db.Tablestyle;
            tb.Position = pt;

            int rows = data.Count + 2;
            int cols = 2;

            tb.SetSize(rows, cols);
            tb.SetRowHeight(rowHeight);
            tb.SetColumnWidth(colWidth);

            tb.Cells.TextHeight = rowHeight / 2;

            // Title
            tb.Cells[0, 0].TextString = "ENTITY COUNT";
            tb.MergeCells(CellRange.Create(tb, 0, 0, 0, 1));
            tb.Cells[0, 0].Alignment = CellAlignment.MiddleCenter;

            //  Headers
            tb.Cells[1, 0].TextString = "Entity";
            tb.Cells[1, 1].TextString = "Count";

            tb.Cells[1, 0].Alignment = CellAlignment.MiddleCenter;
            tb.Cells[1, 1].Alignment = CellAlignment.MiddleCenter;

            //  Data
            int row = 2;
            foreach (var item in data)
            {
                tb.Cells[row, 0].TextString = item.Key;
                tb.Cells[row, 1].TextString = item.Value.ToString();

                tb.Cells[row, 0].Alignment = CellAlignment.MiddleCenter;
                tb.Cells[row, 1].Alignment = CellAlignment.MiddleCenter;

                row++;
            }

            tb.GenerateLayout();

            ms.AppendEntity(tb);
            tr.AddNewlyCreatedDBObject(tb, true);
        }
    }
}