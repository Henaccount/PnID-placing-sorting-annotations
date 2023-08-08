using Autodesk.AutoCAD;

using Autodesk.AutoCAD.Runtime;

using Autodesk.AutoCAD.ApplicationServices;

using Autodesk.AutoCAD.DatabaseServices;

using Autodesk.AutoCAD.EditorInput;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Text.RegularExpressions;
using Autodesk.ProcessPower.PnPCommonDbx;
using Autodesk.ProcessPower.PnIDObjects;
using PlantApp = Autodesk.ProcessPower.PlantInstance.PlantApplication;
using Autodesk.ProcessPower.ProjectManager;
using Autodesk.ProcessPower.DataLinks;

namespace pssCommands
{
    public class EqpAnnotation
    {
        public string tagstrpart { get; set; }
        public string tagintpart { get; set; }
        public string tagendstr { get; set; }
        public string firstandnumber { get; set; }
        public string secondandnumber { get; set; }
        public ObjectId eqpId { get; set; }
    }

    public class Program
    {
        [CommandMethod("EQPANNOPLACE", CommandFlags.UsePickSet)]
        public void PlaceAnnotations()
        {
            string myAnnoStyle = "EqpBomAnno";
            Point3d insertPoint = new Point3d(28,20,0);
            double offset = 0.3557;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            PromptSelectionResult selres = ed.SelectImplied();
            if (selres.Status == PromptStatus.Error) return;

            foreach (ObjectId selobjectid in selres.Value.GetObjectIds())
            {
                // get the style id from the name input

                ObjectId styleId = PnIDStyleUtils.GetStyleIdFromName(myAnnoStyle, true);

                // now create the annotation object

                Annotation anno = new Annotation();

                // and place it

                anno.Create(insertPoint, styleId, selobjectid);

                insertPoint = new Point3d(insertPoint.X, insertPoint.Y - offset, insertPoint.Z);
            }
        }

        [CommandMethod("EQPANNOSORT")]

        public void ListAttributes()
        {
            string configstr = "";
            double originX = 0.0;
            double originY = 0.0;
            double originZ = 0.0;
            double xshift = 0.0;
            double yshift = 0.0;

            Editor ed =

                    Application.DocumentManager.MdiActiveDocument.Editor;

            PromptResult pr = ed.GetString("\nconfiguration string: ");
            if (pr.Status != PromptStatus.OK)
            {
                return;
            }
            else
                configstr = pr.StringResult;

            if (!configstr.Equals(""))
            {
                string[] configArr = configstr.Split(new char[] { ',' });

                foreach (string cstr in configArr)
                {
                    string cstrkey = cstr.Split(new char[] { '=' })[0].Trim();
                    string cstrval = cstr.Split(new char[] { '=' })[1].Trim();

                    switch (cstrkey)
                    {
                        case "originX":
                            originX = Convert.ToDouble(cstrval);
                            break;
                        case "originY":
                            originY = Convert.ToDouble(cstrval);
                            break;
                        case "originZ":
                            originZ = Convert.ToDouble(cstrval);
                            break;
                        case "xshift":
                            xshift = Convert.ToDouble(cstrval);
                            break;
                        case "yshift":
                            yshift = Convert.ToDouble(cstrval);
                            break;
                        default:
                            ed.WriteMessage("\nconfiguration string is not according to the rules!");
                            return;
                            break;
                    }
                }


            }
            else
            {
                ed.WriteMessage("No configuration string was provided\n");
                return;
            }

            List<EqpAnnotation> t = new List<EqpAnnotation>();



            Database db =

              HostApplicationServices.WorkingDatabase;

            Transaction tr =

              db.TransactionManager.StartTransaction();


            // Start the transaction

            try
            {

                // Build a filter list so that only

                // block references are selected

                TypedValue[] filList = new TypedValue[1] {

          new TypedValue((int)DxfCode.Start, "INSERT")

        };

                SelectionFilter filter =

                  new SelectionFilter(filList);

                /*PromptSelectionOptions opts =

                  new PromptSelectionOptions();

                opts.MessageForAdding = "Select block references: ";*/

                PromptSelectionResult res = ed.SelectAll(filter);

                //ed.GetSelection(opts, filter);


                // Do nothing if selection is unsuccessful

                if (res.Status != PromptStatus.OK)

                    return;


                SelectionSet selSet = res.Value;

                ObjectId[] idArray = selSet.GetObjectIds();

                List<ObjectId> forSelection = new List<ObjectId>();

                foreach (ObjectId blkId in idArray)
                {
                    //isannotation?
                    if (!Autodesk.ProcessPower.PnIDObjects.PnIDAnnotationUtils.IsPnIDAnnotation(blkId))
                        continue;

                    BlockReference blkRef =

                      (BlockReference)tr.GetObject(blkId,

                        OpenMode.ForRead);

                    BlockTableRecord btr =

                      (BlockTableRecord)tr.GetObject(

                        blkRef.BlockTableRecord,

                        OpenMode.ForRead

                      );

                    //ed.WriteMessage("\nBlock: " + btr.Name );

                    btr.Dispose();


                    Autodesk.AutoCAD.DatabaseServices.AttributeCollection attCol =

                      blkRef.AttributeCollection;

                    bool isEqpListItem = false;
                    string acttagstr = "";

                    foreach (ObjectId attId in attCol)
                    {

                        AttributeReference attRef =

                          (AttributeReference)tr.GetObject(attId,

                            OpenMode.ForRead);


                        string tagstr = attRef.Tag;

                        string textstr = attRef.TextString;

                        if (tagstr.IndexOf(".Description") != -1)
                            isEqpListItem = true;

                        if (tagstr.IndexOf(".Tag") != -1)
                            acttagstr = textstr;
                        //ed.WriteMessage(str);

                    }

                    if (isEqpListItem)
                    {
                        forSelection.Add(blkId);

                        var actAnno = new EqpAnnotation();
                        Regex regex = new Regex(@"([A-Z]+)(\d+)([A-Z]*)");
                        Match match = regex.Match(acttagstr);
                        if (match.Groups[1].Success && match.Groups[2].Success)
                        {
                            //ed.WriteMessage(match.Groups[1] + "###" + match.Groups[2] + "###" + match.Groups[3]);
                            actAnno.tagstrpart = match.Groups[1].Value;
                            if (actAnno.tagstrpart.Length == 1)
                                actAnno.tagstrpart += "_";
                            actAnno.tagintpart = match.Groups[2].Value;
                            actAnno.tagintpart = String.Format("{0:00000}", Convert.ToInt32(actAnno.tagintpart));
                            if (match.Groups[3].Success)
                            {
                                actAnno.tagendstr = match.Groups[3].Value;
                                if (actAnno.tagendstr.Equals(""))
                                    actAnno.tagendstr = "_";
                            }
                            else
                                actAnno.tagendstr = "_";

                            actAnno.firstandnumber = actAnno.tagstrpart.Substring(0, 1) + actAnno.tagintpart + actAnno.tagendstr;
                            actAnno.secondandnumber = actAnno.tagstrpart.Substring(1, 1) + actAnno.tagintpart + actAnno.tagendstr;

                            //ed.WriteMessage("\n" + actAnno.tagsortable);
                        }
                        else
                            continue;

                        actAnno.eqpId = blkId;
                        t.Add(actAnno);
                    }

                }

                t = t.OrderBy(l => l.firstandnumber).ThenBy(l => l.secondandnumber).ToList();


                Point3d origin = new Point3d(originX, originY, originZ);

                int rowcount = 0;
                int colcount = -1;
                string oldnumber = "";
                string oldfirstchar = "";

                foreach (EqpAnnotation annoObj in t)
                {
                    BlockReference blkRef = (BlockReference)tr.GetObject(annoObj.eqpId, OpenMode.ForWrite);
                    string actnumber = annoObj.tagintpart + annoObj.tagendstr;
                    ed.WriteMessage("\nactnumber " + actnumber);
                    string actfirstchar = annoObj.tagstrpart.Substring(0, 1);
                    if (!actnumber.Equals(oldnumber) || !actfirstchar.Equals(oldfirstchar))
                    {
                        ++colcount;
                        rowcount = 0;
                    }
                    double theX = origin.X + colcount * xshift;
                    double theY = origin.Y - rowcount * yshift;
                    Point3d destination = new Point3d(theX, theY, blkRef.Position.Z);
                    Vector3d vector = blkRef.Position.GetVectorTo(destination);
                    Matrix3d matrix = Matrix3d.Displacement(vector);
                    blkRef.TransformBy(matrix);
                    oldnumber = actnumber;
                    oldfirstchar = actfirstchar;
                    ++rowcount;
                }

                //
                ObjectId[] forSelArr = forSelection.ToArray();
                ed.SetImpliedSelection(forSelArr);

                tr.Commit();

            }

            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {

                ed.WriteMessage(("Exception: " + ex.Message));

            }

            finally
            {

                tr.Dispose();

            }




        }

    }
}
