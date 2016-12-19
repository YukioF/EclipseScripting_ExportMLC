using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;

using System.Windows.Media;
using System.Windows.Media.Media3D;
using System.Windows.Controls;
using System.IO;
using System.Text.RegularExpressions;

using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace VMS.TPS
{

    public class Script
    {
        public Script()
        {
        }

        public void Execute(ScriptContext context /*, System.Windows.Window window*/)
        {
            //**************************
            // Please check the parameters.
            //**************************
            String path_to_OutputFolder = @"C:\Users\Administrator\Desktop";
            double scaling_factor = 1.0;
            //**************************

            Patient patient = context.Patient;
            if (patient == null)
            {
                throw new ApplicationException("Please open a plan or a plan sum before running the script.");
            }

            PlanSetup planSetup = context.PlanSetup;
            PlanSum psum = context.PlanSumsInScope.FirstOrDefault();
            if (planSetup == null && psum == null)
            {
                throw new ApplicationException("Please open a plan or a plan sum before running the script.");
            }
            SelectedPlanningItem = planSetup != null ? (PlanningItem)planSetup : (PlanningItem)psum;
            SelectedStructureSet = planSetup != null ? planSetup.StructureSet : psum.PlanSetups.First().StructureSet;

            // Retrieve StructureSet
            //StructureSet structureSet = planSetup.StructureSet;
            if (SelectedStructureSet == null)
            {
                throw new ApplicationException("The selected plan does not reference a StructureSet.");
            }
            // Retrieve image
            VMS.TPS.Common.Model.API.Image image = SelectedStructureSet.Image;

            //-------------------------------------
            // Plan
            //-------------------------------------

            var beams = planSetup.Beams;

            foreach (var beam in beams)
            {
                if (!beam.IsSetupField)
                {

                    String fileName = path_to_OutputFolder + @"\" + patient.FirstName + patient.LastName + @"_" + patient.Id + @"_" + planSetup.Id + @".txt";
                    StreamWriter writer = new StreamWriter(fileName, false);

                    String machine = beam.TreatmentUnit.Id;

                    //MLC Plan Type
                    int MLCPlanType = (int)beam.MLCPlanType;
                    //MLCPlanType
                    // 1 MLC plan type: DoseDynamic( The gantry does not rotate. )
                    // 2 MLC plan type: ArcDynamic
                    // 3 MLC plan type: VMAT

                    int Count = beam.ControlPoints.Count;
                    // Exporting number of control points
                    writer.WriteLine(Count);

                    var JawPositions = beam.ControlPoints.ElementAt(0).JawPositions;

                    if (MLCPlanType == 1 || MLCPlanType == 2 || MLCPlanType == 3)
                    {
                        double MUw = 0.0;
                        double gantryAngle = 0.0;
                        double collimatorAngle = 0.0;
                        double patientSupportAngle = 0.0;
                        
                        foreach (var controlPoints in beam.ControlPoints)
                        {
                            // Exporting number of control points
                            MUw = controlPoints.MetersetWeight;
                            writer.WriteLine(MUw.ToString("0.0000000"));

                            gantryAngle = controlPoints.GantryAngle;
                            
                            collimatorAngle = controlPoints.CollimatorAngle;
                            
                            patientSupportAngle = controlPoints.PatientSupportAngle;
                            
                            //MLC -----
                            var LeafPositions = controlPoints.LeafPositions;
                            for (int i = 0; i < 60; i++)
                            {
                                double mlc_a = LeafPositions[0, i];
                                mlc_a = mlc_a*scaling_factor;
                                double mlc_b = LeafPositions[1, i];
                                mlc_b = mlc_b*scaling_factor;
                                String mlc_positions = mlc_a.ToString("0.00000") + @", " + mlc_b.ToString("0.00000") + @", 1";
                                // Exporting number of control points
                                writer.WriteLine(mlc_positions);
                            }
                            //MLC ----- end
                        }                     
                    }
                    writer.Close();
                }       
            }

        }

        PlanningItem SelectedPlanningItem { get; set; }
        StructureSet SelectedStructureSet { get; set; }
        Structure SelectedStructure { get; set; }

    }
}
