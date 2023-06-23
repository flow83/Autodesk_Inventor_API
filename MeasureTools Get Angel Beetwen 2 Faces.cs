using Inventor;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using InvAddIn;
using Assembly_Code_CS;


// MeasureTools Get Angel Beetwen 2 Faces

namespace InvAddIn
{
    internal class CommandFunctionButton_05
    {

        public static void CommandFunctionfweButton_05()
        {
            AssemblyDocument oAsmDoc = (AssemblyDocument)Globals.g_inventorApplication.ActiveDocument;
            AssemblyComponentDefinition oAsmCompDef = oAsmDoc.ComponentDefinition;
            TransientObjects oTO = Globals.g_inventorApplication.TransientObjects;
            TransientGeometry oTG = Globals.g_inventorApplication.TransientGeometry;
            TransientBRep oBRep = Globals.g_inventorApplication.TransientBRep;


            ObjectCollection oObjColl = oTO.CreateObjectCollection();
            ObjectCollection oObjColl_Face_Loops = oTO.CreateObjectCollection();
            ObjectCollection oObjColl_Face = oTO.CreateObjectCollection();

            ComponentOccurrences oOcc_Master = oAsmCompDef.Occurrences;


            List<double> oListArea = new List<double>();

            string oLocalization = Microsoft.VisualBasic.Interaction.InputBox("Localization");
            string oCSV_Name = Microsoft.VisualBasic.Interaction.InputBox("oCSV_Name");

            string oPath = oLocalization + @"\" + oCSV_Name + ".csv";
            using (StreamWriter writer = new StreamWriter(oPath))




                foreach (ComponentOccurrence oOcc in oOcc_Master.AllReferencedOccurrences[oAsmCompDef])
                {
                    if (oOcc.DefinitionDocumentType == DocumentTypeEnum.kPartDocumentObject)
                    {
                        PartComponentDefinition oCompDef = (PartComponentDefinition)oOcc.Definition;


                        foreach (SurfaceBody oBody in oCompDef.SurfaceBodies)
                        {
                            oObjColl_Face_Loops.Clear();
                            oListArea.Clear();

                            foreach (Face oFaceBody in oBody.Faces)
                            {

                                if (oFaceBody.EdgeLoops.Count > 1)
                                {

                                    if (oFaceBody.Evaluator.Area > 2)
                                    {
                                        oObjColl_Face_Loops.Add(oFaceBody);
                                        oListArea.Add(oFaceBody.Evaluator.Area);

                                    }

                                }

                            }


                            oListArea.Sort();
                            oObjColl_Face.Clear();

                            foreach (double oItem in oListArea)
                            {


                                foreach (Face oFace in oObjColl_Face_Loops)
                                {
                                    if (oFace.Evaluator.Area == oItem)
                                    {
                                        oObjColl_Face.Add(oFace);
                                    }
                                }
                            }





                            ObjectCollection oObjColl_Face_Top = oTO.CreateObjectCollection();
                            ObjectCollection oObjColl_Face_Bottom = oTO.CreateObjectCollection();

                            oObjColl_Face_Top.Clear();
                            oObjColl_Face_Bottom.Clear();


                            List<double> oListDistance = new List<double>();

                            Face oFace_Top_A = (Face)oObjColl_Face[1];
                            Face oFace_Top_B = (Face)oObjColl_Face[2];
                            Face oFace_Bottom_A = (Face)oObjColl_Face[3];
                            Face oFace_Bottom_B = (Face)oObjColl_Face[4];

                            oObjColl_Face_Top.Add(oFace_Top_A);
                            oObjColl_Face_Top.Add(oFace_Top_B);
                            oObjColl_Face_Bottom.Add(oFace_Bottom_A);
                            oObjColl_Face_Bottom.Add(oFace_Bottom_B);



                            WorkPoint oWorkPoint_Top_A = oCompDef.WorkPoints.AddAtCentroid(oFace_Top_A.EdgeLoops[2], true);
                            WorkPoint oWorkPoint_Top_B = oCompDef.WorkPoints.AddAtCentroid(oFace_Top_B.EdgeLoops[2], true);
                            WorkPoint oWorkPoint_Bottom_A = oCompDef.WorkPoints.AddAtCentroid(oFace_Bottom_A.EdgeLoops[2], true);
                            WorkPoint oWorkPoint_Bottom_B = oCompDef.WorkPoints.AddAtCentroid(oFace_Bottom_B.EdgeLoops[2], true);

                            double oDistance_1_3 = Globals.g_inventorApplication.MeasureTools.GetMinimumDistance(oWorkPoint_Top_A, oWorkPoint_Bottom_A);
                            double oDistance_1_4 = Globals.g_inventorApplication.MeasureTools.GetMinimumDistance(oWorkPoint_Top_A, oWorkPoint_Bottom_B);
                            double oDistance_2_3 = Globals.g_inventorApplication.MeasureTools.GetMinimumDistance(oWorkPoint_Top_B, oWorkPoint_Bottom_A);
                            double oDistance_2_4 = Globals.g_inventorApplication.MeasureTools.GetMinimumDistance(oWorkPoint_Top_B, oWorkPoint_Bottom_B);

                            oListDistance.Add(oDistance_1_3);
                            oListDistance.Add(oDistance_1_4);
                            oListDistance.Add(oDistance_2_3);
                            oListDistance.Add(oDistance_2_4);

                            oListDistance.Sort();

                            double oDistanceMinimum = (double)oListDistance[0];

                            oObjColl.Clear();


                            foreach (Face oFaceTop in oObjColl_Face_Top)
                            {
                                WorkPoint oWorkPoint_A = oCompDef.WorkPoints.AddAtCentroid(oFaceTop.EdgeLoops[2], true);

                                foreach (Face oFaceBottom in oObjColl_Face_Bottom)
                                {
                                    WorkPoint oWorkPoint_B = oCompDef.WorkPoints.AddAtCentroid(oFaceBottom.EdgeLoops[2], true);
                                    double oDistance = Globals.g_inventorApplication.MeasureTools.GetMinimumDistance(oWorkPoint_A, oWorkPoint_B);

                                    if (oDistance == oDistanceMinimum)
                                    {
                                        oObjColl.Add(oFaceTop);
                                        oObjColl.Add(oFaceBottom);
                                        break;
                                    }
                                }

                            }



                            Face oFace_Target_A = (Face)oObjColl[1];
                            Face oFace_Target_B = (Face)oObjColl[2];

                            WorkPlane oWorkPlane_A = (WorkPlane)oCompDef.WorkPlanes.AddByPlaneAndPoint(oFace_Target_A, oFace_Target_A.Vertices[1], true);
                            WorkPlane oWorkPlane_B = (WorkPlane)oCompDef.WorkPlanes.AddByPlaneAndPoint(oFace_Target_B, oFace_Target_B.Vertices[1], true);

                            double oAngel_Radians = oWorkPlane_A.Plane.Normal.AngleTo(oWorkPlane_B.Plane.Normal);
                            double oAngel_Degrees_Reverse_Quad = (oAngel_Radians * 180) / Math.PI;
                            double oAngel_Degrees = (oAngel_Degrees_Reverse_Quad - 180) * -1;

                            ComponentOccurrence oOcc_Parent = oOcc.ParentOccurrence;

                            string oOcc_Parent_Name = oOcc_Parent.Name;
                            string oOcc_Name = oOcc.Name;
                            string oAngel_Degrees_Text = oAngel_Degrees.ToString();

                            string oTextLine = $"{oOcc_Parent_Name}; {oOcc_Name};{oAngel_Degrees_Text}";
                            writer.WriteLine(oTextLine);




                        }


                    }


                }



        }
    }
}