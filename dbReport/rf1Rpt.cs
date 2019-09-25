using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;
using Teigha.DatabaseServices;
using Bricscad.ApplicationServices;
using Bricscad.EditorInput;
using Teigha.Runtime;
//alright here's what we need by tomorrow
//don't bother being clever
//export 1 TBlock_Attributes csv file
//EXPORT 1 Tblock_Properties CSV FILE
//IMPORT these into Richard's CADVALIDATION database
//use whatever ridiculous SQL thing you might need
//EntityChecks structures aren't relevant for now, but going to want to perform a EntityCheck on blockdefinitions which, if passed, passes that blockdefinition into a collection for further EntityChecks? or passes objectIds inside that definition to EntityChecks?

namespace checkDwgs
{
    public class EntityCheck<ruleVal>
    {
        ruleVal rule;
        string property;
        string comparison;
        object[] compareParams;
        Dictionary<ObjectId, int> space;
        Dictionary<PropertyInfo, int> layer;
        int total;
        List<ObjectId> errVal = new List<ObjectId>();
        bool reverseMethod;
        public void Check(ObjectId checkObj)
        {
            using (DBObject thisObj = checkObj.GetObject(OpenMode.ForRead, false))
            {
                ObjectId thisSpace = thisObj.OwnerId;
                Type thisType = thisObj.GetType();
                PropertyInfo thisLayer, checkVal;
                checkVal = thisType.GetProperty(property);
                //alright, so, c# doesn't like reflection with generics
                //right now that makes some compromises necessary
                //we're using this overload: http://msdn.microsoft.com/en-us/library/6hy0h0z1.aspx
                //HEY RICHARD I DON'T THINK I UNDERSTAND REFLECTION WITH GENERICS AT AAAALLLLLLLLL
                try
                {
                    MethodInfo thisComparer = typeof(ruleVal).GetMethod(comparison, new[] { typeof(ruleVal) });
                    if (reverseMethod ^ (bool)thisComparer.Invoke(checkVal.GetValue(thisObj), compareParams)) return;
                    ++total;
                    if (!errVal.Contains(checkObj)) errVal.Add(checkObj);
                    if (space.ContainsKey(thisSpace)) ++space[thisSpace]; else space.Add(thisSpace, 1);
                    if (thisType.GetProperties().Where(prop => prop.Name == "Layer").Count() > 0)
                    {
                        if (layer.ContainsKey(thisLayer = thisType.GetProperty("Layer"))) ++layer[thisLayer]; else layer.Add(thisLayer, 1);
                    }
                }
                catch (ArgumentNullException)
                {
                    try
                    {
                        MethodInfo[] allMethods = typeof(ruleVal).GetMethods().Where(method => method.Name == comparison).ToArray();
                        foreach (MethodInfo thisComparer in allMethods)
                        {
                            try
                            {
                                if (reverseMethod ^ (bool)thisComparer.Invoke(checkVal.GetValue(thisObj), compareParams)) return;
                                ++total;
                                if (!errVal.Contains(checkObj)) errVal.Add(checkObj);
                                if (space.ContainsKey(thisSpace)) ++space[thisSpace]; else space.Add(thisSpace, 1);
                                if (thisType.GetProperties().Where(prop => prop.Name == "Layer").Count() > 0)
                                {
                                    if (layer.ContainsKey(thisLayer = thisType.GetProperty("Layer"))) ++layer[thisLayer]; else layer.Add(thisLayer, 1);
                                }
                                break;
                            }
                            catch (System.Exception)
                            {
                                //dunno what to put in here
                                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Exception of some sort");
                            }

                        }
                    }
                    catch (System.Exception)
                    {
                        //dunno what to put in here
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Argument Null Exception unrecoverable");
                    }
                }
                catch (System.Reflection.AmbiguousMatchException)
                {
                    try
                    {
                        MethodInfo[] allMethods = typeof(ruleVal).GetMethods().Where(method => method.Name == comparison && method.DeclaringType == typeof(ruleVal)).ToArray();
                        if (reverseMethod ^ (bool)allMethods[0].Invoke(checkVal.GetValue(thisObj), compareParams)) return;
                        ++total;
                        if (!errVal.Contains(checkObj)) errVal.Add(checkObj);
                        if (space.ContainsKey(thisSpace)) ++space[thisSpace]; else space.Add(thisSpace, 1);
                        if (thisType.GetProperties().Where(prop => prop.Name == "Layer").Count() > 0)
                        {
                            if (layer.ContainsKey(thisLayer = thisType.GetProperty("Layer"))) ++layer[thisLayer]; else layer.Add(thisLayer, 1);
                        }
                    }
                    catch (System.Exception)
                    {
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("Ambiguous Match Exception did not work on first seemingly acceptable value");
                    }
                }
            }
        }
        public EntityCheck(string propertyName, ruleVal rv, string comparisonMethod)
        {
            rule = rv; property = propertyName; space = new Dictionary<ObjectId, int>(); layer = new Dictionary<PropertyInfo, int>();
            total = 0; comparison = comparisonMethod; compareParams = new object[] { rv }; reverseMethod = true;
        }
        public EntityCheck(string propertyName, ruleVal rv, string comparisonMethod, bool defaultFalse)
        {
            rule = rv; property = propertyName; space = new Dictionary<ObjectId, int>(); layer = new Dictionary<PropertyInfo, int>();
            total = 0; reverseMethod = defaultFalse; comparison = comparisonMethod; compareParams = new object[] { rv };
        }
        public EntityCheck(string propertyName, ruleVal rv, string comparisonMethod, bool defaultFalse, object[] compareParameters)
        {
            rule = rv; property = propertyName; space = new Dictionary<ObjectId, int>(); layer = new Dictionary<PropertyInfo, int>();
            total = 0; reverseMethod = defaultFalse; comparison = comparisonMethod; compareParams = compareParameters;
        }

    }
    public class TableDefCheck<ruleVal>
    {
        ruleVal rule;
        string property;
        List<ObjectId> subentities;
        List<PropertyInfo> errVal;
        bool reverseMethod;
        bool Check(ObjectId checkObj)
        {
            DBObject thisObj = checkObj.GetObject(OpenMode.ForRead, false);
            ObjectId thisSpace = thisObj.OwnerId;
            PropertyInfo checkVal;
            if ((checkVal = thisObj.GetType().GetProperty(property)).Equals(rule)) return false;
            if (checkVal.Equals(null)) return false;
            return true;
        }
        public TableDefCheck(string propertyName, ruleVal rv, bool defaultFalse)
        { rule = rv; property = propertyName; reverseMethod = defaultFalse; }

    }
    public class Rf1Rpt
    {

        [CommandMethod("CheckList")]
        static public void ExamineDwgs()
        {
            StreamWriter alldwgscsv = new StreamWriter("G:\\VERIFICATION\\RF1\\ALLDWGS.CSV", append: false);
            StreamWriter threedobjectcsv = new StreamWriter("G:\\VERIFICATION\\RF1\\3DOBJECTS.CSV", append: false);
            StreamWriter badtxtcsv = new StreamWriter("G:\\VERIFICATION\\RF1\\BADTXT.CSV", append: false);
            StreamWriter blocksinblockscsv = new StreamWriter("G:\\VERIFICATION\\RF1\\BLOCKSINBLOCKS.CSV", append: false);
            StreamWriter intangiblescsv = new StreamWriter("G:\\VERIFICATION\\RF1\\INTANGIBLES.CSV", append: false);
            StreamWriter layerrorscsv = new StreamWriter("G:\\VERIFICATION\\RF1\\LAYERRORS.CSV", append: false);
            StreamWriter mtxtcsv = new StreamWriter("G:\\VERIFICATION\\RF1\\MTXT.CSV", append: false);
            StreamWriter purgeablescsv = new StreamWriter("G:\\VERIFICATION\\RF1\\PURGEABLES.CSV", append: false);
            StreamWriter rederrscsv = new StreamWriter("G:\\VERIFICATION\\RF1\\REDERRS.CSV", append: false);
            StreamWriter revblockscsv = new StreamWriter("G:\\VERIFICATION\\RF1\\REVBLOCKS.CSV", append: false);
            StreamWriter stylerrorscsv = new StreamWriter("G:\\VERIFICATION\\RF1\\STYLERRORS.CSV", append: false);
            StreamWriter tablescsv = new StreamWriter("G:\\VERIFICATION\\RF1\\TABLES.CSV", append: false);
            //StreamWriter titleblockscsv = new StreamWriter("G:\\VERIFICATION\\RF1\\TITLEBLOCKS.CSV", append: false);
            StreamWriter tbattscsv = new StreamWriter("G:\\VERIFICATION\\RF1\\TBlock_Attributes.CSV", append: false);
            StreamWriter tbpropscsv = new StreamWriter("G:\\VERIFICATION\\RF1\\TBlock_Properties.CSV", append: false);
            StreamWriter underlayscsv = new StreamWriter("G:\\VERIFICATION\\RF1\\UNDERLAY.CSV", append: false);
            StreamWriter viewportscsv = new StreamWriter("G:\\VERIFICATION\\RF1\\VIEWPORTS.CSV", append: false);
            alldwgscsv.WriteLine("Full Path to Drawing | Filename | Path | Layouts");
            threedobjectcsv.WriteLine("Full Path to Drawing | 3D Objects | in layout | in layer | in block definition");
            badtxtcsv.WriteLine("Full Path to Drawing | Removable text");
            blocksinblockscsv.WriteLine("Full Path to Drawing | Blocks-within-blocks");
            intangiblescsv.WriteLine("Full Path to Drawing | Hidden Dynamic Block Objects | in layout | in layer");
            layerrorscsv.WriteLine("Full Path to Drawing | Non-standard layernames | Non-standard layer properties");
            mtxtcsv.WriteLine("Full Path to Drawing | MTEXT objects | in layout | in layer | in block definition");
            purgeablescsv.WriteLine("Full Path to Drawing | Purgeable layers | Purgeable blocks | Purgeable text styles");
            rederrscsv.WriteLine("Full Path to Drawing | Red Objects | in layout | in layer | in block definition");
            revblockscsv.WriteLine("Full Path to Drawing | REV | DATE | MOC# | DRFTR | DSGNR | ENGNR | LEADR");
            stylerrorscsv.WriteLine("Full Path to Drawing | Non-MonoTXT fonts");
            tablescsv.WriteLine("Full Path to Drawing | Table Objects | in layout | in layer | in block definition");
            //TBLOCK_Attributes contains: DWGPATH | TBLOCK_NUMBER 
            tbattscsv.WriteLine("Full Path to Drawing | TBlock Number | DWGNUMBER | REVISION | SHEETNO | DISC | DWG_DESCR1 | DWG_DESCR2 | DWG_DESCR3 | DWG_DESCR4 | INITS | SAVEDATE");
            tbpropscsv.WriteLine("Full Path to Drawing | TBlock Number | Insertion | Definition | Layout | Layername | Block Scale | Attributes | Irregularities");
            underlayscsv.WriteLine("Full Path to Drawing | DWG XREFs | Image/Model Underlays | Excel Datalinks");
            viewportscsv.WriteLine("Full Path to Drawing | Clipped/Rotated VPort | VP Locked? | VP Layer");
            void DwgCheck(Database thisFile)
            {
                Transaction tr = thisFile.TransactionManager.StartTransaction();
                using (tr)
                {
                    string[] acceptableTBNames = { "B_RBLOCK", "B_TBLOCK", "D_RBLOCK", "D_TBLOCK", "C_RBLOCK", "F_RBLOCK", "C_TBLOCK", "A_RBLOCK", "A_TBLOCK", "H_RBLOCK", "H_TBLOCK", "F_TBLOCK", "F_PBLOCK" };
                    ObjectIdCollection allLayouts = new ObjectIdCollection();
                    ObjectIdCollection tbDefs = new ObjectIdCollection();
                    ObjectIdCollection usedBlocks = new ObjectIdCollection();
                    ObjectIdCollection titleBlocks = new ObjectIdCollection();
                    ObjectIdCollection usedLayers = new ObjectIdCollection();
                    ObjectIdCollection usedStyles = new ObjectIdCollection();
                    ObjectIdCollection vPorts = new ObjectIdCollection();
                    ObjectIdCollection txtObjects = new ObjectIdCollection();
                    int i = 1;
                    BlockTable blockDefs = (BlockTable)thisFile.BlockTableId.GetObject(OpenMode.ForRead, false);
                    EntityCheck<int> rederrs = new EntityCheck<int>("ColorIndex", 1, "Equals");
                    EntityCheck<string> layerrs = new EntityCheck<string>("Layer", "", "Equals", true);
                    Type[] typeList = { typeof(BlockReference), typeof(MText), typeof(Solid), typeof(PolyFaceMesh), typeof(PolygonMesh) };
                    foreach (ObjectId thisBlockId in blockDefs)
                    {
                        
                        BlockTableRecord thisBlock = (BlockTableRecord)thisBlockId.GetObject(OpenMode.ForRead, false);
                        if (thisBlock.IsLayout)
                        {
                            foreach (ObjectId thisEntity in thisBlock)
                            {
                                rederrs.Check(thisEntity);
                                layerrs.Check(thisEntity);
                                DBObject thisObject = thisEntity.GetObject(OpenMode.ForRead, false);
                                if (thisObject.GetType() == typeList[0])
                                {
                                    if (acceptableTBNames.Contains(((BlockReference)thisObject).Name.ToUpper()))
                                        titleBlocks.Add(thisEntity);
                                }
                            }
                        }
                        else if (acceptableTBNames.Contains(thisBlock.Name.ToUpper()))
                        {
                            tbDefs.Add(thisBlockId);
                        }
                    }
                    //dbUmmm.SimpleUpdateDB("Server=NTSRVR22;Database=CADValidation;User Id=cad_reporter;Password=pa$$word;", "INSERT INTO DWG_Files (DWGPath) VALUES ('" + thisFile.Filename + "')");
                    if (titleBlocks.Count == 0)
                    {
                        tbpropscsv.WriteLine(thisFile.Filename + "| NO TITLEBLOCK FOUND | | | | | | | | |");
                        tbattscsv.WriteLine(thisFile.Filename + "| NO TITLEBLOCK FOUND | | | | | | | | |");
                    }
                    else foreach (ObjectId thisTBlock in titleBlocks)
                        {
                            string tbIrregularity = "";
                            BlockReference thisTBlockReference = (BlockReference)thisTBlock.GetObject(OpenMode.ForRead, false);
                            string[] tbAttributeNames = { "DWGNUMBER", "REVISION", "SHEETNO", "DISC", "DWG_DESCR1", "DWG_DESCR2", "DWG_DESCR3", "DWG_DESCR4", "INITS", "SAVEDATE" };
                            string[] tbAttributes = { "NO DWG NUMBER FOUND", "NO REV NUMBER FOUND", "NO SHEET NUMBER FOUND", "NO DISCIPLINE ASSIGNED", "NO DWG TYPE ASSIGNED", "NO INFO", "NO PLANT NAME ASSIGNED", "NO ABU ASSIGNED", "NO PLOT STAMP INITIALS", "NO SAVE DATE" };
                            foreach (ObjectId thisAttribute in thisTBlockReference.AttributeCollection)
                            {
                                AttributeReference thisAttributeReference = (AttributeReference)thisAttribute.GetObject(OpenMode.ForRead, false);
                                if (tbAttributeNames.Contains(thisAttributeReference.Tag)) tbAttributes[Array.IndexOf(tbAttributeNames, thisAttributeReference.Tag)] = thisAttributeReference.TextString;
                            }
                            if (thisTBlockReference.Name.ToUpper()[2] == 'T')
                            {

                                if (!tbAttributes[1].Equals("NO REV NUMBER FOUND") || !tbAttributes[2].Equals("NO SHEET NUMBER FOUND"))
                                {
                                    tbIrregularity = "Mutant titleblock detected";
                                }
                            }
                            if (thisTBlockReference.Name.ToUpper()[2] == 'R')
                            {
                                if (tbAttributes[2].Equals("-") || tbAttributes[2].Equals(".") || tbAttributes[2].Equals(" ") || tbAttributes[2].Equals("") || tbAttributes[2].Equals("X") ||
                                    tbAttributes[1].Equals("-") || tbAttributes[1].Equals(".") || tbAttributes[1].Equals(" ") || tbAttributes[1].Equals("") || tbAttributes[1].Equals("X"))
                                {
                                    tbIrregularity = "Irregular Rev/Sht Number";
                                }
                                if (tbAttributes[1].Equals("NO REV NUMBER FOUND") || tbAttributes[2].Equals("NO SHEET NUMBER FOUND"))
                                {
                                    tbIrregularity = "Mutant titleblock detected";
                                }
                            }
                            string[] TBlockPropertyNames = { "DWGPath", "TBlockNumber", "Position", "Definition", "LayoutName", "Layer", "ScaleFactors", "Attributes", "Irregularity" };
                            string[] TBlockPropertyValues = new string[TBlockPropertyNames.Length];
                            TBlockPropertyValues[0] = thisFile.Filename;
                            TBlockPropertyValues[1] = titleBlocks.IndexOf(thisTBlock).ToString();
                            TBlockPropertyValues[2] = thisTBlockReference.Position.ToString();
                            TBlockPropertyValues[3] = thisTBlockReference.Name;
                            TBlockPropertyValues[4] = ((Layout)(((BlockTableRecord)thisTBlockReference.OwnerId.GetObject(OpenMode.ForRead, false)).LayoutId.GetObject(OpenMode.ForRead, false))).LayoutName;
                            TBlockPropertyValues[5] = thisTBlockReference.Layer;
                            TBlockPropertyValues[6] = thisTBlockReference.ScaleFactors.ToString();
                            TBlockPropertyValues[7] = thisTBlockReference.AttributeCollection.Count.ToString();
                            TBlockPropertyValues[8] = tbIrregularity;
                            StringBuilder qryText = new StringBuilder();
                            qryText.Append("INSERT INTO TBlock_Properties VALUES (");
                            int iterator;
                            for (iterator = 0; iterator < TBlockPropertyNames.Length; iterator++ )
                            {
                               qryText.Append("'"+ TBlockPropertyValues[iterator] + "', ");
                            }
                            qryText.Remove(qryText.Length - 3, 2);
                            qryText.Append("')");
                            tbpropscsv.WriteLine(TBlockPropertyValues[0] + "|" + TBlockPropertyValues[1] + "|" + TBlockPropertyValues[2] + "|" + TBlockPropertyValues[3] + "|" + TBlockPropertyValues[4] + "|" + TBlockPropertyValues[5] + "|" + TBlockPropertyValues[6] + "|" + TBlockPropertyValues[7] + "|" + TBlockPropertyValues[8]);
                            tbattscsv.WriteLine(thisFile.Filename + "|" + titleBlocks.IndexOf(thisTBlock) + "|" + tbAttributes[0] + "|" + tbAttributes[1] + "|" + tbAttributes[2]  + "|" + tbAttributes[3] + "|" + tbAttributes[4] + "|" + tbAttributes[5] + "|" + tbAttributes[6] + "|" + tbAttributes[7] + "|" + tbAttributes[8] + "|" + tbAttributes[9]);
                            //dbUmmm.SimpleUpdateDB("Server=ntsrvr22;Database=CADValidation;User Id=cad_reporter;Password=pa$$word;", qryText.ToString());
                            qryText.Clear();
                            qryText.Append("INSERT INTO TBlock_Attributes VALUES ( '" + TBlockPropertyValues[0] + "', '" + TBlockPropertyValues[1] + "', ");
                            for (iterator = 0; iterator < tbAttributeNames.Length; iterator++)
                            {
                                //qryText.Append(tbAttributeNames[iterator] + " " + TBlockPropertyValues[iterator]);
                                qryText.Append("'" + tbAttributes[iterator] + "', ");
                            }
                            qryText.Remove(qryText.Length - 3, 2);
                            qryText.Append("')");
                            //dbUmmm.SimpleUpdateDB("Server=ntsrvr22;Database=CADValidation;User Id=cad_reporter;Password=pa$$word;", qryText.ToString());
                        }

                }
                tr.Dispose();
            }
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            PromptResult feed = ed.GetString("\nEnter the full name (including path) of a .txt format list of checkable DWG files:  ");
            StreamReader dwg = new StreamReader(feed.StringResult);
            if (feed.Status == PromptStatus.OK)
            {
                
                while (!dwg.EndOfStream)
                {
                    using (Database db = new Database(false, true))
                    {
                        String dwgpath = dwg.ReadLine();
                        try
                        {
                            db.ReadDwgFile(dwgpath, FileShare.Read, true, "");
                        }
                        catch (System.Exception)
                        {
                            ed.WriteMessage("\nUnable to read file " + dwgpath + "\n");
                            return;
                        }
                        DwgCheck(db);
                        db.CloseInput(true);
                    }
                }
            }
            alldwgscsv.Close();
            threedobjectcsv.Close();
            badtxtcsv.Close();
            blocksinblockscsv.Close();
            intangiblescsv.Close();
            layerrorscsv.Close();
            mtxtcsv.Close();
            purgeablescsv.Close();
            rederrscsv.Close();
            revblockscsv.Close();
            stylerrorscsv.Close();
            tablescsv.Close();
            tbattscsv.Close();
            tbpropscsv.Close();
            underlayscsv.Close();
            viewportscsv.Close();
        }
    }
}