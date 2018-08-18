using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using Newtonsoft.Json;
using System.Collections;
using System.Diagnostics;
using FormSerialisation;
using BrightIdeasSoftware;
using Jint;
using FileHelpers;
using WindowsFormsApplication1;

namespace StatureTool
{

    public partial class Form1 : System.Windows.Forms.Form
    {

        private VMS.TPS.Common.Model.API.Application app;
        private string UserID;

        //needed to track ascending vs descending sorting of listview
        private int mrnSortColumn = -1;
        private int matrixSortColumn = -1;
        private int matchesSortColumn = -1;

        List<string> MRNList = new List<string>();
        List<string> archivedMRNs = new List<string>();
        List<string> MD5List = new List<string>();
        List<string> StructureList = new List<string>();

        //string OutputDir = @"C:\temp\StructureNameAnalyser\results";

        public class StructureInfo
        {
            public string Id { set; get; }
            public bool isEmpty { set; get; }
            public double vol { set; get; }
            public double d95 { set; get; }
            public string d95unit { set; get; }
            public double dMean { set; get; }
            public string dMeanUnit { set; get; }
            public double dMedian { set; get; }
            public string dMedianUnit { set; get; }
        }

        public class PlanPathAndStructures
        {
            public string MRN { set; get; }
            public string Course { set; get; }
            public DateTime? CourseStart { set; get; }
            public string Plan { set; get; }
            public double Dose { set; get; }
            public string DoseUnit { set; get; }
            public int? Fractions { set; get; }
            public List<StructureInfo> Structures = new List<StructureInfo>();
        }

        List<PlanPathAndStructures> matchedPlans = new List<PlanPathAndStructures>();

        //Separate Retrieval and Filter Mechanism by storing all relevant data locally using the following connected objects
        public class LocalStructure
        {
            public string Id { set; get; }
            public bool IsEmpty { set; get; }
            public double vol { set; get; }
            public double d95 { set; get; }
            public string d95unit { set; get; }
            public double dMean { set; get; }
            public string dMeanUnit { set; get; }
            public double dMedian { set; get; }
            public string dMedianUnit { set; get; }
        }
        public class LocalPlan
        {
            public string Id { set; get; }
            public bool IsTreated { set; get; }
            public bool IsValid { set; get; }
            public double Dose { set; get; }
            public string DoseUnit { set; get; }
            public int? Fractions { set; get; }
            public List<LocalStructure> Structures = new List<LocalStructure>();
        }
        public class LocalCourse
        {
            public string Id { set; get; }
            public DateTime? StartDateTime { set; get; }
            public List<LocalPlan> Plans = new List<LocalPlan>();
        }
        public class LocalPatient
        {
            public string mrn { set; get; }
            public List<LocalCourse> Courses = new List<LocalCourse>();
        }

        public List<LocalPatient> localPatients = new List<LocalPatient>();

        public class FaEntry
        {
            public FaEntry(int occurence, int empty, double volCum, double volMin, double volMax, int dmeanNan, double dmeanCum, double dmeanMin, double dmeanMax, int d95Nan, double d95Cum, double d95Min, double d95Max, int volMatch, int doseMatch)
            {
                this.occurence = occurence;
                this.empty = empty;
                this.volCum = volCum;
                this.volMin = volMin;
                this.volMax = volMax;
                this.dmeanNan = dmeanNan;
                this.dmeanCum = dmeanCum;
                this.dmeanMin = dmeanMin;
                this.dmeanMax = dmeanMax;
                this.d95Nan = d95Nan;
                this.d95Cum = d95Cum;
                this.d95Min = d95Min;
                this.d95Max = d95Max;
                this.volMatch = volMatch;
                this.doseMatch = doseMatch;
            }

            public int occurence { set; get; }
            public int empty { set; get; }
            public Double volCum { set; get; }
            public Double volMin { set; get; }
            public Double volMax { set; get; }
            public int dmeanNan { set; get; }
            public Double dmeanCum { set; get; }
            public Double dmeanMin { set; get; }
            public Double dmeanMax { set; get; }
            public int d95Nan { set; get; }
            public Double d95Cum { set; get; }
            public Double d95Min { set; get; }
            public Double d95Max { set; get; }
            public int volMatch { set; get; }
            public int doseMatch { set; get; }
        }

        public class FaFilter
        {
            public FaFilter(string name)
            {
                this.name = name;
            }
            public FaFilter(string name, string regEx)
            {
                this.name = name;
                this.regEx = regEx;
            }
            [JsonConstructor]
            public FaFilter(string name, int frequencyTreshold, string regEx, string doseType, string doseMin, string doseMax, string doseTreshold, string volMin, string volMax, string volTreshold) : this(name)
            {
                this.frequencyTreshold = frequencyTreshold;
                this.regEx = regEx;
                this.doseType = doseType;
                this.doseMin = doseMin;
                this.doseMax = doseMax;
                this.doseTreshold = doseTreshold;
                this.volMin = volMin;
                this.volMax = volMax;
                this.volTreshold = volTreshold;
            }


            //using string for everything except frequencyThreshold
            public string name { set; get; }
            public int frequencyTreshold { set; get; }
            public string regEx { set; get; }
            public string doseType { set; get; }
            public string doseMin { set; get; }
            public string doseMax { set; get; }
            public string doseTreshold { set; get; }
            public string volMin { set; get; }
            public string volMax { set; get; }
            public string volTreshold { set; get; }
        }

        public List<FaFilter> faFilters = new List<FaFilter>();


        //Synonym Template classes

        [DelimitedRecord("\t")]
        public class OLV_Template
        {
            public string semClass { get; set; }
            public string group { get; set; }
            public string synonyms { get; set; }
            public OLV_Template() { }
            public OLV_Template(string semClass, string group, string synonyms)
            {
                this.semClass = semClass;
                this.group = group;
                this.synonyms = synonyms;
            }
        }
        public class SemClass
        {
            public string name { get; set; }
            public string group { get; set; }
            public HashSet<string> synonyms;

            public SemClass()
            {
                this.synonyms = new HashSet<string>();
            }
            public SemClass(string name, string group, HashSet<string> synonyms)
            {
                this.name = name;
                this.group = group;
                this.synonyms = synonyms;
            }
            public string SynonymCommaString()
            {
                string output = "";
                foreach (string synonym in this.synonyms)
                {
                    output += synonym + ", ";
                }
                output = output.TrimEnd(',', ' '); //remove last ", "
                return (output);
            }
        }
        public class Template  //for mapping structure names to their semantic class
        {
            private HashSet<string> allMembers; //contains names of all semantic classes and their synonyms 
            public HashSet<string> allMembersStem; //contains names of all semantic classes and their synonyms (for synonyms only "stems" before curly braces are added)
            private Dictionary<string, SemClass> semClasses; //associates a semantic class a set of synonyms

            public Template()
            {
                this.allMembers = new HashSet<string>();
                this.allMembersStem = new HashSet<string>(); //contains all synonym_stems (used for highlighting)
                this.semClasses = new Dictionary<string, SemClass>();
            }

            public int AddSynonymToClass(string semClassName, string synonym)
            {
                //extract stem before "{"
                string synonym_stem = "";
                int curl_pos = synonym.IndexOf('{');
                if (curl_pos > 0)
                {
                    synonym_stem = synonym.Substring(0, curl_pos);
                }
                else
                {
                    synonym_stem = synonym;
                }

                if (!this.allMembers.Contains(synonym))
                {
                    SemClass semClass;
                    if (this.semClasses.TryGetValue(semClassName, out semClass))
                    {
                        if (semClass.synonyms.Add(synonym))
                        {
                            this.allMembers.Add(synonym);
                            if (!this.allMembersStem.Contains(synonym_stem)) { this.allMembersStem.Add(synonym_stem); }
                            return 0;
                        }
                        else { return 1; } //synonym already exists, should never happen as caught by first if
                    }
                    else
                    {
                        return 2; //semClass does not exist
                    }
                }
                else return 1; //synonym already exists as a name somewhere in template
            }
            public int AddSemClass(SemClass semClass)
            {
                if (!this.allMembers.Contains(semClass.name))
                {
                    HashSet<string> tempSet = new HashSet<string>(this.allMembers);
                    tempSet.IntersectWith(semClass.synonyms);
                    if (tempSet.Count == 0)
                    {
                        this.semClasses.Add(semClass.name, semClass); //add synonyms to semClasses dictionary
                        this.allMembers.Add(semClass.name);
                        this.allMembersStem.Add(semClass.name);
                        this.allMembers.UnionWith(semClass.synonyms); //add syonyms to allMembers HashSet
                        foreach (string synonym in semClass.synonyms) //add stems if not already there
                        {
                            //extract stem before "{"
                            string synonym_stem = "";
                            int curl_pos = synonym.IndexOf('{');
                            if (curl_pos > 0)
                            {
                                synonym_stem = synonym.Substring(0, curl_pos);
                            }
                            else
                            {
                                synonym_stem = synonym;
                            }
                            if (!this.allMembersStem.Contains(synonym_stem)) { this.allMembersStem.Add(synonym_stem); }
                        }
                        return 0;
                    }
                    else { return 2; }  //at least one synonym in synonyms exists as name somewhere in template

                }
                else { return 1; }  //semClassName already exists as name somewhere in template
            }
            public string ContentAsString()
            {
                string output = "";
                foreach (KeyValuePair<string, SemClass> kvp in this.semClasses)
                {
                    output += kvp.Key + "\t" + kvp.Value.group + "\t";  //could have also used kvp.Value.name instead of kvp.Key
                    output += kvp.Value.SynonymCommaString();
                    output += Environment.NewLine;
                }
                return output;
            }
            public int LoadFromString(string multiLineString)
            {
                Template temp_template = new Template(); //creates a temporary second template object test whether can add 

                var lines = multiLineString.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length == 0) { return 3; }  // problem with multiline string 

                foreach (string line in lines)
                {
                    var tabs = line.Split(new string[] { "\t" }, StringSplitOptions.None);
                    if (tabs.Length > 1)
                    {
                        SemClass temp_semClass = new SemClass();
                        temp_semClass.synonyms = new HashSet<string>();

                        if (tabs[0] != "") { temp_semClass.name = tabs[0].Trim(); } else { return 3; }  //no LSSN set
                        temp_semClass.group = tabs[1].Trim();
                        if (tabs.Length > 2)
                        {
                            var temp_synonyms = tabs[2].Split(new string[] { "," }, StringSplitOptions.None);
                            foreach (string temp_synonym in temp_synonyms)
                            {
                                temp_semClass.synonyms.Add(temp_synonym.Trim());
                            }
                            int temp_return;
                            temp_return = temp_template.AddSemClass(temp_semClass);
                            if (temp_return != 0) { return temp_return; } //pass on error code
                        }
                        else
                        {
                            int temp_return;
                            temp_return = temp_template.AddSemClass(temp_semClass);
                            if (temp_return != 0) { return temp_return; } //pass on error code
                        }
                    }
                    else { return 3; }  // problem with multiline string   
                }
                // copy temp to real template
                this.allMembers = temp_template.allMembers;
                this.allMembersStem = temp_template.allMembersStem;
                this.semClasses = temp_template.semClasses;
                return 0;
            }

            public void ContentToCSV(string path_to_file)  //";" delimeted
            {
                List<OLV_Template> list = new List<OLV_Template>();
                foreach (KeyValuePair<string, SemClass> kvp in this.semClasses)
                {
                    string synonym_string = "";
                    foreach (string synonym in kvp.Value.synonyms)
                    {
                        synonym_string += synonym + ", ";
                    }
                    synonym_string = synonym_string.TrimEnd(',', ' '); //remove last ", "
                    list.Add(new OLV_Template(kvp.Value.name, kvp.Value.group, synonym_string));
                }

                OLV_Template[] array = list.ToArray();
                var engine = new FileHelperEngine<OLV_Template>();
                engine.WriteFile(path_to_file, array);
            }

            public void LoadFromCSV(string path_to_file)  //";" delimeted and "," delimited (synonyms)
            {
                var engine = new FileHelperEngine<OLV_Template>();
                var result = engine.ReadFile(path_to_file);

                if (result.Length > 0)
                {
                    this.allMembers.Clear();
                    this.allMembersStem.Clear();
                    this.semClasses.Clear();

                    foreach (OLV_Template record in result)
                    {
                        HashSet<string> tempHashSet = new HashSet<string>();
                        string[] synonyms = record.synonyms.Split(',');
                        for (int i = 0; i < synonyms.Length; i++)
                        {
                            tempHashSet.Add(synonyms[i].Trim());
                        }
                        if (tempHashSet.Count > 0 && record.semClass != "")
                        {
                            this.AddSemClass(new SemClass(record.semClass, record.group, tempHashSet));
                        }
                    }
                }
            }

            public List<OLV_Template> ContentAsListForOLV()
            {
                List<OLV_Template> output = new List<OLV_Template>();
                //foreach (KeyValuePair<string, HashSet<string>> semClass in this.semClasses)
                foreach (KeyValuePair<string, SemClass> kvp in this.semClasses)
                {

                    output.Add(new OLV_Template(kvp.Value.name, kvp.Value.group, kvp.Value.SynonymCommaString()));
                }
                return output;
            }
        }

        private Template buildingTemplate = new Template();

        // OLV Models
        private List<OLV_MappingDetails> MappingDetailsOLVmodel = new List<OLV_MappingDetails>();
        //maybe requires similar one for List<OLV_Template>  (see just above)

        public Form1()
        {
            InitializeComponent();

            //EnableButton(Login_Button);
            //DisableButton(SelectMRNList_Button);
            //DisableButton(DVHExport_Button);
            //System.IO.Directory.CreateDirectory(OutputDir); // Create temp directory if not already exists

        }


        [DebuggerStepThrough()]
        private void pullDosimetricData()
        {
            int ptMax;
            int cMax;
            int plMax;
            int stMax;
            double prog;

            TreeNode archivedParentNode;

            localPatients.Clear();
            matchedPlans.Clear();
            valid_treeView.Nodes.Clear();
            review_treeView.Nodes.Clear();
            none_treeView.Nodes.Clear();
            archived_treeView.Nodes.Clear();
            Matches_listView.Items.Clear();
            indicatorValid_label.Visible = false;
            indicatorReview_label.Visible = false;
            indicatorNone_label.Visible = false;
            indicatorArchived_label.Visible = false;
            indicatorLocalData_label.Visible = false;

            archivedParentNode = new TreeNode("ARCHIVED - plan data not available due to archiving");

            List<LocalPatient> exceptionPatients = new List<LocalPatient>();
            string exception_log = "";

            Console.WriteLine("=== PULL DATA LOCALLY (Start) ===");
            ptMax = MRN_listView.Items.Count;
            if (ptMax == 0) { Console.WriteLine("No MRN list loaded."); }

            for (int ptIterator = 0; ptIterator < ptMax; ptIterator++)
            {
                //Reset iterator if offset chosen
                if (ptIterator == 0)
                {
                    if (mrnOffset_numericUpDown.Value <= MRN_listView.Items.Count)
                    {
                        ptIterator = (int)mrnOffset_numericUpDown.Value - 1;
                    }
                }

                double ptProg = (double)ptIterator / ptMax;
                prog = ptProg;
                Console.WriteLine("Progress: " + prog.ToString("P") + "(Pt #" + (ptIterator + 1).ToString() + " of " + ptMax + ")");

                string MRN = MRN_listView.Items[ptIterator].Text;

                //string overriddenCourseID = "";
                //string overriddenPlanID = "";
                //if (MRN_listView.Items[ptIterator].SubItems.Count > 2) //both course and plan override set
                //{
                //    overriddenCourseID = MRN_listView.Items[ptIterator].SubItems[1].Text;
                //    overriddenPlanID = MRN_listView.Items[ptIterator].SubItems[2].Text;
                //}
                //bool overridden_flag = false;
                //if (overriddenCourseID != "" && overriddenPlanID != "") { overridden_flag = true; }

                bool archived_flag = false;
                bool in_archived_list_flag = false;
                if (archivedMRNs.Contains(MRN))
                {
                    archived_flag = true;
                    in_archived_list_flag = true;
                }
                else
                {
                    if (checkDir_checkBox.Checked)
                    {
                        if (!Directory.Exists(@"\\rnsroncprod\VA_Data$\filedata\Patients\" + MRN))
                        {
                            archived_flag = true;
                            if (MRN.StartsWith("0"))
                            {
                                if (Directory.Exists(@"\\rnsroncprod\VA_Data$\filedata\Patients\" + MRN.TrimStart('0')))
                                {
                                    archived_flag = false;
                                }
                            }
                        }
                    }
                }


                if (archived_flag)
                {
                    if (in_archived_list_flag)
                    {
                        archivedParentNode.Nodes.Add(MRN, "Pt " + (ptIterator + 1).ToString() + " (" + MRN + ") => via ARCHIVED LIST");
                        Console.WriteLine("ARCHIVED (via archived black list)");
                    }
                    else
                    {
                        archivedParentNode.Nodes.Add(MRN, "Pt " + (ptIterator + 1).ToString() + " (" + MRN + ") => via DIRECTORY CHECK");
                        Console.WriteLine("ARCHIVED (via directory check)");
                    }
                    continue; //skip this patient.
                }


                LocalPatient localPatient = new LocalPatient();
                Course c;
                LocalCourse localCourse;
                string temp_log = "";
                //pull all relevant data for current patient from TPS
                try
                {
                    Patient pt = app.OpenPatientById(MRN);

                    localPatient.mrn = MRN;
                    temp_log = MRN;

                    //If ?corrupt data structure this is where ESAPI usually throws exception  
                    cMax = pt.Courses.Count();


                    temp_log += ", Courses (n= " + cMax + ")";

                    for (int cIterator = 0; cIterator < cMax; cIterator++)
                    {

                        double cProg = (double)cIterator / (ptMax * cMax);
                        prog = ptProg + cProg;
                        Console.WriteLine("Progress: " + prog.ToString("P"));

                        c = pt.Courses.ToList()[cIterator];
                        temp_log += "C" + cIterator.ToString() + ", ";
                        localCourse = new LocalCourse();
                        localCourse.Id = c.Id;
                        localCourse.StartDateTime = c.StartDateTime;
                        localPatient.Courses.Add(localCourse);

                        plMax = c.PlanSetups.Count();
                        temp_log += ", Plans (n= " + plMax + ")";


                        for (int plIterator = 0; plIterator < plMax; plIterator++)
                        {
                            double plProg = (double)plIterator / (ptMax * cMax * plMax);
                            prog = ptProg + cProg + plProg;
                            Console.WriteLine("Progress: " + prog.ToString("P"));

                            PlanSetup pl = c.PlanSetups.ToList()[plIterator];
                            temp_log += "P" + plIterator.ToString() + ", ";

                            LocalPlan localPlan = new LocalPlan();
                            localPlan.Id = pl.Id;
                            localPlan.IsTreated = pl.IsTreated;
                            localPlan.IsValid = pl.IsDoseValid;
                            localPlan.Dose = pl.TotalPrescribedDose.Dose;
                            localPlan.DoseUnit = pl.TotalPrescribedDose.UnitAsString;
                            localPlan.Fractions = pl.UniqueFractionation.NumberOfFractions;

                            localPatient.Courses[cIterator].Plans.Add(localPlan);

                            if (pl.StructureSet != null) //needed to avoid structureless plans when relaxing isDoseValid and is isTreated conditions
                            {
                                StructureSet ss = pl.StructureSet;
                                DoseValue dv;
                                DVHData dvhData;
                                stMax = ss.Structures.Count();
                                temp_log += ", Structures (n= " + stMax + ")";

                                if (stMax > 0)
                                {
                                    for (int stIterator = 0; stIterator < stMax; stIterator++)
                                    {
                                        Structure st = ss.Structures.ToList()[stIterator];
                                        if (st.DicomType != "MARKER")  //could potentially lead to no added structures despite structure count > 0 
                                        {
                                            temp_log += ", S" + stIterator.ToString();
                                            LocalStructure localStructure = new LocalStructure();
                                            localStructure.Id = st.Id;
                                            localStructure.IsEmpty = st.IsEmpty;
                                            localStructure.vol = st.Volume;
                                            if (preFetch_checkBox.Checked)
                                            {
                                                dv = pl.GetDoseAtVolume(st, 95, VolumePresentation.Relative, DoseValuePresentation.Absolute);
                                                localStructure.d95 = dv.Dose;
                                                localStructure.d95unit = dv.UnitAsString;
                                                dvhData = pl.GetDVHCumulativeData(st, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, 0.1);
                                                if (dvhData != null)
                                                {
                                                    dv = dvhData.MeanDose;
                                                    localStructure.dMean = dv.Dose;
                                                    localStructure.dMeanUnit = dv.UnitAsString;
                                                    dv = dvhData.MedianDose;
                                                    localStructure.dMedian = dv.Dose;
                                                    localStructure.dMedianUnit = dv.UnitAsString;
                                                }
                                            }
                                            localPatient.Courses[cIterator].Plans[plIterator].Structures.Add(localStructure);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    localPatients.Add(localPatient);
                }
                catch (Exception e)
                {
                    //attempt to catch retrieval errors - unfortunately doesn't work for archived patients => blacklist solution
                    //MessageBox.Show(e.Message, "Error during data retrival from TPS", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    exception_log += temp_log + Environment.NewLine;
                    exceptionPatients.Add(localPatient);
                    archivedParentNode.Nodes.Add(MRN, "Pt " + (ptIterator + 1).ToString() + " (" + MRN + ") => via API EXCEPTION");
                    Console.WriteLine("ESAPI EXCEPTION (?corrupt file)");
                    //DEBUG: write pt objects to file if exception called 
                    //using (StreamWriter file = File.CreateText(System.Windows.Forms.Application.StartupPath + @"\pull_exception_pt_objs\" + ptIterator.ToString() +".json"))
                    //{
                    //    JsonSerializer serializer = new JsonSerializer();
                    //    serializer.Serialize(file, pt);
                    //}
                }
                finally
                {
                    app.ClosePatient();
                }
            }

            if (exception_log != "")
            {
                File.WriteAllText(System.Windows.Forms.Application.StartupPath + @"\pull_exceptions.txt", exception_log);
            }
            //if (exceptionPatients.Count > 0)
            //{
            //    using (StreamWriter file = File.CreateText(System.Windows.Forms.Application.StartupPath + @"\pull_exceptions.json"))
            //    {
            //        JsonSerializer serializer = new JsonSerializer();
            //        serializer.Serialize(file, exceptionPatients);
            //    }
            //}
            indicatorArchived_label.Text = archivedParentNode.Nodes.Count.ToString() + " archived";
            indicatorArchived_label.Visible = true;
            indicatorLocalData_label.Text = localPatients.Count.ToString() + " loaded locally";
            indicatorLocalData_label.Visible = true;
            if (archivedParentNode.Nodes.Count > 0)
            {
                archived_treeView.Nodes.Add(archivedParentNode);
            }

            Console.WriteLine("Progress: 100.00 %");
            Console.WriteLine("=== PULL DATA LOCALLY (Finish) ===");
        }

        private void filterDosimetricData()
        {
            int ptMax;
            int cMax;
            int plMax;
            int stMax;
            double prog;

            TreeNode validParentNode;
            TreeNode reviewParentNode;
            TreeNode noneParentNode;
            //TreeNode archivedParentNode;
            TreeNode ptNode;
            TreeNode cNodes;
            TreeNode cNode;
            TreeNode plNodes;
            TreeNode plNode;
            TreeNode infoNode;
            TreeNode exclCodesNode;
            TreeNode inclCodesNode;
            TreeNode stNodes;

            List<PlanPathAndStructures> clinicalPlanCandidates = new List<PlanPathAndStructures>();
            List<StructureInfo> structureInfoContainer = new List<StructureInfo>();
            bool courseIsMatch;
            bool planIsMatch;

            bool tumourSpecificIsMatch;
            int BrCaLT_int = 0;
            int BrCaRT_int = 0;
            int BrCaBL_int = 0;
            int BrCaBR_int = 0;
            int BrCaCW_int = 0;
            int BrCaSCF_int = 0;
            int BrCaAX_int = 0;
            int BrCaIMC_int = 0;

            string excludeCode; //?TODO handle with array to capture multiple exclusion codes per plan
            Dictionary<string, List<string>> filterCodes = new Dictionary<string, List<string>>();
            filterCodes.Add("pLevel", new List<string>());  //used to show overrides and inclusions (second round)
            filterCodes.Add("cLevel", new List<string>());
            filterCodes.Add("plLevel", new List<string>());



            matchedPlans.Clear();
            valid_treeView.Nodes.Clear();
            review_treeView.Nodes.Clear();
            none_treeView.Nodes.Clear();
            //archived_treeView.Nodes.Clear();

            //reset Matches listview including turning sorting off to avoid slowing down filter run
            Matches_listView.Items.Clear();
            Matches_listView.Sorting = SortOrder.None;
            Matches_listView.ListViewItemSorter = null;
            foreach (ColumnHeader ch in Matches_listView.Columns)
            {
                if (ch.Name.StartsWith("brCa")) { Matches_listView.Columns.Remove(ch); }
            }
            Matches_listView.BeginUpdate();

            if (tumourSpecific_checkBox.Checked)
            {
                if (brCancer_radioButton.Checked)
                {
                    Matches_listView.Columns.Add("brCa_LT", "LT");
                    Matches_listView.Columns.Add("brCa_RT", "RT");
                    Matches_listView.Columns.Add("brCa_BL", "BL");
                    Matches_listView.Columns.Add("brCa_Lat", "Lat check");
                    Matches_listView.Columns.Add("brCa_BR", "BR");
                    Matches_listView.Columns.Add("brCa_CW", "CW");
                    Matches_listView.Columns.Add("brCa_Loco", "Loco check");
                    Matches_listView.Columns.Add("brCa_SCF", "SCF");
                    Matches_listView.Columns.Add("brCa_AX", "AX");
                    Matches_listView.Columns.Add("brCa_IMC", "IMC");
                    Matches_listView.Columns.Add("brCa_Regional", "Regional check");
                }
            }

            indicatorValid_label.Visible = false;
            indicatorReview_label.Visible = false;
            indicatorNone_label.Visible = false;
            //indicatorArchived_label.Visible = false;

            Console.WriteLine("=== FILTER LOCAL DATA (Start) ===");
            ptMax = localPatients.Count;
            if (ptMax == 0) { Console.WriteLine("No data locally available. Pull first."); }

            validParentNode = new TreeNode("VALID - successful filtering");
            reviewParentNode = new TreeNode("REVIEW NEEDED - use override mechanism or tighten filter criteria");
            noneParentNode = new TreeNode("NONE - consider loosening filter criteria");
            //archivedParentNode = new TreeNode("ARCHIVED - plan data not available due to archiving");

            for (int ptIterator = 0; ptIterator < ptMax; ptIterator++)
            {
                clinicalPlanCandidates.Clear();
                // Empty store of filter codes at the beginning of each patient
                filterCodes["pLevel"].Clear();
                filterCodes["cLevel"].Clear();
                filterCodes["plLevel"].Clear();



                double ptProg = (double)ptIterator / ptMax;
                prog = ptProg;
                Console.WriteLine("Progress: " + prog.ToString("P") + "(Pt #" + (ptIterator + 1).ToString() + " of " + ptMax + ")");

                LocalPatient pt = localPatients[ptIterator];
                ptNode = new TreeNode("Pt " + (ptIterator + 1).ToString() + " (" + pt.mrn + ")");
                ptNode.Name = pt.mrn;  //put a name so, can later be used for navigation

                string overriddenCourseID = "";
                string overriddenPlanID = "";
                ListViewItem override_lvi = MRN_listView.Items.Find(pt.mrn, false).FirstOrDefault();
                if (override_lvi != null)
                {
                    if (override_lvi.SubItems.Count > 2) //both course and plan override set
                    {
                        overriddenCourseID = override_lvi.SubItems[1].Text;
                        overriddenPlanID = override_lvi.SubItems[2].Text;
                    }
                }
                bool overridden_flag = false;
                if (overriddenCourseID != "" && overriddenPlanID != "") { overridden_flag = true; }

                cMax = pt.Courses.Count();
                cNodes = new TreeNode("Courses (n=" + cMax.ToString() + ")");
                ptNode.Nodes.Add(cNodes);

                for (int cIterator = 0; cIterator < cMax; cIterator++)
                {
                    //reset isCourseMatch flag and exclusionCode
                    courseIsMatch = true;
                    excludeCode = "";
                    filterCodes["cLevel"].Clear();

                    double cProg = (double)cIterator / (ptMax * cMax);
                    prog = ptProg + cProg;
                    Console.WriteLine("Progress: " + prog.ToString("P"));

                    LocalCourse c = pt.Courses[cIterator];
                    if (!c.StartDateTime.Value.IsInRange(From_dateTimePicker.Value, To_dateTimePicker.Value)) { courseIsMatch = false; excludeCode = " => [2a]"; filterCodes["cLevel"].Add("2a"); };
                    if (!c.Id.StartsWith(CourseMatchStart_TextBox.Text)) { courseIsMatch = false; excludeCode = " => [3a]"; filterCodes["cLevel"].Add("3a"); };
                    if (CourseMatchExcludes_textBox.Lines.Any(s => c.Id.Contains(s))) { courseIsMatch = false; excludeCode = " => [3b]"; filterCodes["cLevel"].Add("3b"); };

                    if (advFilters_checkBox.Checked && courseRegex_textBox.Text != "")
                    {
                        Regex courseExclRegex = new Regex(@courseRegex_textBox.Text);
                        if (courseExclRegex.IsMatch(c.Id)) { courseIsMatch = false; excludeCode = " => [3c]"; filterCodes["cLevel"].Add("3c"); };

                    }

                    if (advFilters_checkBox.Checked && courseRegexIncl_textBox.Text != "")
                    {
                        Regex courseInclRegex = new Regex(@courseRegexIncl_textBox.Text);
                        if (!courseInclRegex.IsMatch(c.Id)) { courseIsMatch = false; excludeCode = " => [3d]"; filterCodes["cLevel"].Add("3d"); };
                    }

                    // Tumour specific matching (at course level)

                    if (tumourSpecific_checkBox.Checked)
                    {
                        tumourSpecificIsMatch = true;
                        string regex_tumour_specific = "";

                        if (brCancer_radioButton.Checked)
                        {
                            if (brCaLaterality_checkBox.Checked)
                            {
                                if (brCaLatLeft_radioButton.Checked)
                                { regex_tumour_specific = "(?=.*" + brCaLatLTregex_textBox.Text + ")"; }
                                if (brCaLatRight_radioButton.Checked)
                                { regex_tumour_specific = "(?=.*" + brCaLatRTregex_textBox.Text + ")"; }
                                if (brCaLatBilat_radioButton.Checked)
                                { regex_tumour_specific = "(?=.*" + brCaLatBLregex_textBox.Text + ")"; }
                            }
                            if (brCaRegionBR_checkBox.Checked)
                            { regex_tumour_specific = regex_tumour_specific + "(?=.*" + brCaRegionBRregex_textBox.Text + ")"; }
                            if (brCaRegionCW_checkBox.Checked)
                            { regex_tumour_specific = regex_tumour_specific + "(?=.*" + brCaRegionCWregex_textBox.Text + ")"; }
                            if (brCaRegionSCF_checkBox.Checked)
                            { regex_tumour_specific = regex_tumour_specific + "(?=.*" + brCaRegionSCFregex_textBox.Text + ")"; }
                            if (brCaRegionAX_checkBox.Checked)
                            { regex_tumour_specific = regex_tumour_specific + "(?=.*" + brCaRegionAXregex_textBox.Text + ")"; }
                            if (brCaRegionIMC_checkBox.Checked)
                            { regex_tumour_specific = regex_tumour_specific + "(?=.*" + brCaRegionIMCregex_textBox.Text + ")"; }
                        }

                        regex_tumour_specific = ".*" + regex_tumour_specific + ".*";
                        Regex tumourSpecificInclRegex = new Regex(@regex_tumour_specific, RegexOptions.IgnoreCase);
                        if (!tumourSpecificInclRegex.IsMatch(c.Id)) { tumourSpecificIsMatch = false; excludeCode = " => [7a]"; filterCodes["cLevel"].Add("7a"); };
                    }
                    else
                    { tumourSpecificIsMatch = true; } // set true so always matches if no tumour specific filtering


                    cNode = cNodes.Nodes.Add(c.Id, c.Id + " (Start date: " + String.Format("{0:d}", c.StartDateTime) + ")");

                    plMax = c.Plans.Count();
                    plNodes = cNode.Nodes.Add("Plans (n=" + plMax.ToString() + ")");

                    for (int plIterator = 0; plIterator < plMax; plIterator++)
                    {
                        double plProg = (double)plIterator / (ptMax * cMax * plMax);
                        prog = ptProg + cProg + plProg;
                        Console.WriteLine("Progress: " + prog.ToString("P"));

                        //reset isPlanMatch flag and exclusionCode
                        planIsMatch = true;
                        excludeCode = "";
                        filterCodes["plLevel"].Clear();

                        LocalPlan pl = c.Plans[plIterator];
                        Regex digit_digit = new Regex(@"\d+_\d+");
                        Regex digitdashdigit = new Regex(@"\d+-\d+");
                        Regex fdigit_fdigit = new Regex(@"f\d+_f\d+");
                        if (replanIdentification_checkBox.Checked)
                        {
                            if (digit_digit.IsMatch(pl.Id) || digitdashdigit.IsMatch(pl.Id) || fdigit_fdigit.IsMatch(pl.Id) || pl.Id.EndsWith(":1")) { planIsMatch = false; excludeCode = " => [4a]"; filterCodes["plLevel"].Add("4a"); };
                        }

                        if (advFilters_checkBox.Checked && planRegex_textBox.Text != "")
                        {
                            Regex planExclRegex = new Regex(@planRegex_textBox.Text);
                            if (planExclRegex.IsMatch(pl.Id)) { planIsMatch = false; excludeCode = " => [4b]"; filterCodes["plLevel"].Add("4b"); };
                        }

                        if (advFilters_checkBox.Checked && planRegexIncl_textBox.Text != "")
                        {
                            Regex planInclRegex = new Regex(@planRegexIncl_textBox.Text);
                            if (!planInclRegex.IsMatch(pl.Id)) { planIsMatch = false; excludeCode = " => [4c]"; filterCodes["plLevel"].Add("4c"); };
                        }

                        plNode = plNodes.Nodes.Add(pl.Id, pl.Id + " (" + String.Format("{0:0.00}", pl.Dose.ToString()) + " " + pl.DoseUnit + " in " + pl.Fractions.ToString() + " fractions)");

                        //reset excludeCode
                        excludeCode = "";

                        //whether to take isDoseValid into account (can rule out some old plans WITH STRUCTURES AND DOSE ?old dose calc algorithm) 
                        if (validDose_checkBox.Checked)
                        {
                            if (!pl.IsValid)
                            {
                                planIsMatch = false;
                                excludeCode = " => [5a]";
                                filterCodes["plLevel"].Add("5a");
                            }
                        }

                        if (treatedPlans_checkBox.Checked)
                        {
                            if (!pl.IsTreated) //might need similar approach via is_treated_flag than isDoseValid above
                            {
                                planIsMatch = false;
                                excludeCode = " => [5b]";
                                filterCodes["plLevel"].Add("5b");
                            }
                        }
                        //Exclude plans that match "exclusion" fractionations
                        foreach (string item in ExclDose_listBox.Items)
                        {
                            string td = ""; //can be range ('-')
                            double td_min_in_cgy = 0; //also includes "leniancy" correction
                            double td_max_in_cgy = 0; //also includes "leniancy" correction
                            string fractions = "";
                            int fractions_min = 0;
                            int fractions_max = 0;
                            
                            if (item.Contains('/'))
                            {
                                td_min_in_cgy = Convert.ToDouble(item.Split('/')[0].Trim()) *100 - Convert.ToDouble(ExclDose_numericUpDown.Value);
                                td_max_in_cgy = Convert.ToDouble(item.Split('/')[0].Trim()) * 100 + Convert.ToDouble(ExclDose_numericUpDown.Value);
                                fractions_min = Convert.ToInt32(item.Split('/')[1].Trim());
                                fractions_max = Convert.ToInt32(item.Split('/')[1].Trim());
                            }

                            if (item.Contains("#"))
                            {
                                fractions = item.Split('#')[0].Trim();
                                td_min_in_cgy = 0; //set td range that includes all plans based on these values
                                td_max_in_cgy = 100000; //set td range that includes all plans based on these values

                                if (fractions.Contains("-"))
                                {
                                    fractions_min = Convert.ToInt32(fractions.Split('-')[0].Trim());
                                    fractions_max = Convert.ToInt32(fractions.Split('-')[1].Trim());
                                }
                                else
                                {
                                    fractions_min = Convert.ToInt32(fractions.Trim());
                                    fractions_max = Convert.ToInt32(fractions.Trim());
                                }
                            }

                            if (item.Contains("Gy"))
                            {
                                td = item.Split('G')[0].Trim();
                                fractions_min = 0; //set fraction range that includes all plans based on these values
                                fractions_max = 100; //set fraction range that includes all plans based on these values

                                if (td.Contains("-"))
                                {
                                    td_min_in_cgy = Convert.ToDouble(td.Split('-')[0].Trim()) * 100 - Convert.ToDouble(ExclDose_numericUpDown.Value);
                                    td_max_in_cgy = Convert.ToDouble(td.Split('-')[1].Trim()) * 100 + Convert.ToDouble(ExclDose_numericUpDown.Value);
                                }
                                else
                                {
                                    td_min_in_cgy = Convert.ToDouble(td.Trim()) * 100 - Convert.ToDouble(ExclDose_numericUpDown.Value);
                                    td_max_in_cgy = Convert.ToDouble(td.Trim()) * 100 + Convert.ToDouble(ExclDose_numericUpDown.Value);
                                }
                            }

                            if (pl.Dose >= td_min_in_cgy && pl.Dose <= td_max_in_cgy && pl.Fractions >= fractions_min && pl.Fractions <= fractions_max)
                            {
                                planIsMatch = false;
                                excludeCode = " => [5c]";
                                filterCodes["plLevel"].Add("5c");
                                break; //no need to look further
                            }
                        }


                        //Exclude other plans than match "inclusion" fractionations

                        bool local_incldose_flag = false;  // to be set to true current plans fractionation matches one of incl fractionations
                        foreach (string item in InclDose_listBox.Items)
                        {
                            string td = ""; //can be range ('-')
                            double td_min_in_cgy = 0; //also includes "leniancy" correction
                            double td_max_in_cgy = 0; //also includes "leniancy" correction
                            string fractions = "";
                            int fractions_min = 0;
                            int fractions_max = 0;

                            if (item.Contains("/"))
                            {
                                td_min_in_cgy = Convert.ToDouble(item.Split('/')[0].Trim()) * 100 - Convert.ToDouble(InclDose_numericUpDown.Value);
                                td_max_in_cgy = Convert.ToDouble(item.Split('/')[0].Trim()) * 100 + Convert.ToDouble(InclDose_numericUpDown.Value);
                                fractions_min = Convert.ToInt32(item.Split('/')[1].Trim());
                                fractions_max = Convert.ToInt32(item.Split('/')[1].Trim());
                            }

                            if (item.Contains("#"))
                            {
                                fractions = item.Split('#')[0].Trim();
                                td_min_in_cgy = 0; //set td range that includes all plans based on these values
                                td_max_in_cgy = 100000; //set td range that includes all plans based on these values

                                if (fractions.Contains("-"))
                                {
                                    fractions_min = Convert.ToInt32(fractions.Split('-')[0].Trim());
                                    fractions_max = Convert.ToInt32(fractions.Split('-')[1].Trim());
                                }
                                else
                                {
                                    fractions_min = Convert.ToInt32(fractions.Trim());
                                    fractions_max = Convert.ToInt32(fractions.Trim());
                                }
                            }

                            if (item.Contains("Gy"))
                            {
                                td = item.Split('G')[0].Trim();
                                fractions_min = 0; //set fraction range that includes all plans based on these values
                                fractions_max = 100; //set fraction range that includes all plans based on these values

                                if (td.Contains("-"))
                                {
                                    td_min_in_cgy = Convert.ToDouble(td.Split('-')[0].Trim()) * 100 - Convert.ToDouble(InclDose_numericUpDown.Value);
                                    td_max_in_cgy = Convert.ToDouble(td.Split('-')[1].Trim()) * 100 + Convert.ToDouble(InclDose_numericUpDown.Value);
                                }
                                else
                                {
                                    td_min_in_cgy = Convert.ToDouble(td.Trim()) * 100 - Convert.ToDouble(InclDose_numericUpDown.Value);
                                    td_max_in_cgy = Convert.ToDouble(td.Trim()) * 100 + Convert.ToDouble(InclDose_numericUpDown.Value);
                                }
                            }

                            if (pl.Dose >= td_min_in_cgy && pl.Dose <= td_max_in_cgy && pl.Fractions >= fractions_min && pl.Fractions <= fractions_max)
                            {
                                local_incldose_flag = true;
                                break; //no need to look further
                            }
                        }
                        if (InclDose_listBox.Items.Count > 0 && !local_incldose_flag)  //disregard if no inclusion fractionations set 
                        {
                            planIsMatch = false;
                            excludeCode = " => [5d]";
                            filterCodes["plLevel"].Add("5d");
                        }

                        if (pl.Structures != null) //needed to avoid structureless plans when relaxing isDoseValid and is isTreated conditions
                        {
                            stMax = pl.Structures.Count();
                            stNodes = plNode.Nodes.Add("Structures (n=" + stMax.ToString() + ")");

                            if (stMax > 0)
                            {
                                structureInfoContainer.Clear();
                                for (int stIterator = 0; stIterator < stMax; stIterator++)
                                {
                                    LocalStructure st = pl.Structures[stIterator];
                                    structureInfoContainer.Add(new StructureInfo { Id = st.Id, isEmpty = st.IsEmpty, vol = st.vol, d95 = st.d95, d95unit = st.d95unit, dMean = st.dMean, dMedian = st.dMedian});
                                    stNodes.Nodes.Add(st.Id + " (vol: " + String.Format("{0:0.00}", st.vol) + " cc; d95: " + String.Format("{0:0.00}", st.d95) + " " + st.d95unit + "; d(mean): " + String.Format("{0:0.00}", st.dMean) + " " + st.dMeanUnit + "; d(median): " + String.Format("{0:0.00}", st.dMedian) + " " + st.dMedianUnit + ")");
                                }
                            }
                            else { planIsMatch = false; excludeCode = " => [6a]"; filterCodes["plLevel"].Add("6a"); }
                        }
                        else { planIsMatch = false; excludeCode = " => [6a]"; filterCodes["plLevel"].Add("6a"); }


                        if (overridden_flag)
                        {

                            if (overriddenCourseID == c.Id && overriddenPlanID == pl.Id)  //if overridden values are valid always add as candidats
                            {
                                if (structureInfoContainer.Count() > 0)
                                {

                                    clinicalPlanCandidates.Add(new PlanPathAndStructures { MRN = pt.mrn, Course = c.Id, CourseStart = c.StartDateTime, Plan = pl.Id, Dose = pl.Dose, DoseUnit = pl.DoseUnit, Fractions = pl.Fractions, Structures = new List<StructureInfo>(structureInfoContainer) });
                                    courseIsMatch = false; //set to make sure doesn't get added twice in next block
                                    filterCodes["pLevel"].Add("!! 1a !!");
                                }
                                else
                                {
                                    MessageBox.Show("Overriden plan with MRN '" + pt.mrn + "' does not have structures attached.", "No attached structures", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                                }
                            }

                        }



                        if (courseIsMatch && planIsMatch && tumourSpecificIsMatch)  // if matches filter criteria
                        {
                            clinicalPlanCandidates.Add(new PlanPathAndStructures { MRN = pt.mrn, Course = c.Id, CourseStart = c.StartDateTime, Plan = pl.Id, Dose = pl.Dose, DoseUnit = pl.DoseUnit, Fractions = pl.Fractions, Structures = new List<StructureInfo>(structureInfoContainer) });
                        }


                        infoNode = plNode.Nodes.Add("Plan Meta Info" + excludeCode);
                        //infoNode.Nodes.Add("Type: " + pl.PlanType);
                        //infoNode.Nodes.Add("Intent: " + pl.PlanIntent);
                        //infoNode.Nodes.Add("Planning Approver: " + pl.PlanningApprover);
                        //infoNode.Nodes.Add("Tx Approver: " + pl.TreatmentApprover);
                        infoNode.Nodes.Add("ValidDose Flag: " + pl.IsValid.ToString());
                        infoNode.Nodes.Add("Treated Flag: " + pl.IsTreated.ToString());

                        exclCodesNode = plNode.Nodes.Add("Filter Codes - first round (exclusion)");
                        exclCodesNode.Nodes.Add("Course level: " + string.Join(", ", filterCodes["cLevel"].ToArray()));
                        exclCodesNode.Nodes.Add("Plan level: " + string.Join(", ", filterCodes["plLevel"].ToArray()));


                        //if (courseIsMatch && planIsMatch)  //some physics plans caused crash, if ongoing issue consider to remove check for this type of meta data
                        //{
                        //    if (pl.VerifiedPlan != null)
                        //    {
                        //        infoNode.Nodes.Add("Verified Plan (?self): " + pl.VerifiedPlan.Id);
                        //    }
                        //    else
                        //    {
                        //        infoNode.Nodes.Add("Verified Plan (?self): <null>");
                        //    }
                        //}
                        //else
                        //{
                        //    infoNode.Nodes.Add("Verified Plan (?self): not checked");
                        //}

                        if (pl.Structures != null)
                        {
                            infoNode.Nodes.Add("Attached structures (count): " + pl.Structures.Count().ToString());
                        }
                        else
                        {
                            infoNode.Nodes.Add("Attached structures (count): <null>");
                        }

                        if (filterCodes["plLevel"].Count > 0)
                        { plNode.Text += " => [ " + string.Join(", ", filterCodes["plLevel"].ToArray()) + " ]"; }
                    }

                    if (filterCodes["cLevel"].Count > 0)
                    { cNode.Text += " => [ " + string.Join(", ", filterCodes["cLevel"].ToArray()) + " ]"; }
                    
                }



                if (earliestCourseDate_checkBox.Checked)
                {
                    if (clinicalPlanCandidates.Count() > 0)
                    {
                        //Second round filtering
                        List<PlanPathAndStructures> earliestCourseCandidates = new List<PlanPathAndStructures>();

                        DateTime? minCourseStartDate = clinicalPlanCandidates.Min(x => x.CourseStart);
                        earliestCourseCandidates.Clear();
                        earliestCourseCandidates = clinicalPlanCandidates.Where(x => x.CourseStart == minCourseStartDate).ToList();
                        //set inclusion marker ' => [++ 8a ++]' in tree
                        //foreach (PlanPathAndStructures second_round_match in earliestCourseCandidates)
                        //{
                        //    ptNode.Nodes[0].Nodes[second_round_match.Course].Text += " => [++ 8a ++]";
                        //    //ptNode.Nodes[0].Nodes[second_round_match.Course].Nodes[0].Nodes[second_round_match.Plan].Text += " => [++ 8a ++]";
                        //}

                        if (earliestCourseCandidates.Count() > 0)
                        {
                            ptNode.Nodes[0].Nodes[earliestCourseCandidates.First().Course].Text += " => [++ 8a ++]";
                            filterCodes["pLevel"].Add("++ 8a ++");
                            //ptNode.Nodes[0].Nodes[second_round_match.Course].Nodes[0].Nodes[second_round_match.Plan].Text += " => [++ 8a ++]";
                        }
                        clinicalPlanCandidates = earliestCourseCandidates;
                    }
                }

                if (sameNumberStructures_checkBox.Checked)  //can only be selected if restricted to earliest course (filter 8a above)
                {
                    if (clinicalPlanCandidates.Count() > 0)
                    {
                        //read out of current 
                        //List<PlanPathAndStructures> temp_results_grouped_by_course = new List<PlanPathAndStructures>();

                        //temp_results_2nd_round_filtering.Add(clinicalPlanCandidates.First());

                        List<PlanPathAndStructures> FirstInGroupSameCourseSameStrucCountCandidates = new List<PlanPathAndStructures>();

                        bool same_struc_number_flag = false;
                        var temp_results_grouped_by_course = clinicalPlanCandidates.GroupBy(x => x.Course);
                        foreach (var group in temp_results_grouped_by_course)
                        {
                            List<PlanPathAndStructures> candidatesGroup = group.ToList();
                            if (sameStructureNames_radioButton.Checked)
                            {
                                if (candidatesGroup.Count() > 0)
                                {
                                    bool same_structure_names_flag = true;
                                    List<StructureInfo> firstList = candidatesGroup.First().Structures;
                                    List<StructureInfo> secondList;
                                    foreach (PlanPathAndStructures candidate in candidatesGroup)
                                    {
                                        secondList = candidate.Structures;
                                        //TODO
                                        bool areEqual = Enumerable.SequenceEqual(firstList.OrderBy(fElement => fElement), secondList.OrderBy(sElement => sElement));
                                        if (!areEqual) { same_structure_names_flag = false; }
                                    }
                                    if (same_structure_names_flag)
                                    {
                                        foreach (PlanPathAndStructures candidate in candidatesGroup)
                                        {
                                            //ptNode.Nodes[0].Nodes[candidate.Course].Text += " => [++ 9a ++]";
                                            //ptNode.Nodes[0].Nodes[candidate.Course].Nodes[0].Nodes[candidate.Plan].Text += " => [++ 9a ++]";
                                            ptNode.Nodes[0].Nodes[candidate.Course].Nodes[0].Nodes[candidate.Plan].Nodes[0].Text += " => [++ 9a ++]";
                                            filterCodes["pLevel"].Add("++ 9a ++");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                int minStrucCount = candidatesGroup.Min(x => x.Structures.Count);
                                int maxStrucCount = candidatesGroup.Max(x => x.Structures.Count);
                                if (minStrucCount == maxStrucCount && minStrucCount > 0)
                                {
                                    if (candidatesGroup.Count() > 1)
                                    {
                                        FirstInGroupSameCourseSameStrucCountCandidates.Add(candidatesGroup.First());
                                        same_struc_number_flag = true;
                                        foreach (PlanPathAndStructures candidate in candidatesGroup)
                                        {
                                            //ptNode.Nodes[0].Nodes[candidate.Course].Text += " => [++ 9a ++]";
                                            //ptNode.Nodes[0].Nodes[candidate.Course].Nodes[0].Nodes[candidate.Plan].Text += " => [++ 9a ++]";
                                            ptNode.Nodes[0].Nodes[candidate.Course].Nodes[0].Nodes[candidate.Plan].Nodes[0].Text += " => [++ 9a ++]";
                                            filterCodes["pLevel"].Add("++ 9a ++");
                                        }
                                    }
                                }
                            }


                        }
                        if (same_struc_number_flag)
                        { clinicalPlanCandidates = FirstInGroupSameCourseSameStrucCountCandidates; }
                    }
                }


                if (filterCodes["pLevel"].Count > 0)
                {
                    ptNode.Text += " => [ " + string.Join(", ", filterCodes["pLevel"].ToArray()) + " ]";
                }




                //TODO if (overridden_flag and is_valid_override) {} else { switch... => correct place to put override

                switch (clinicalPlanCandidates.Count)
                {
                    case 1:
                        //ptNode.Nodes.Add("SINGLE APPROPRIATE PLAN - successful filtering");
                        validParentNode.Nodes.Add(ptNode);
                        Console.WriteLine("VALID - SINGLE APPROPRIATE PLAN");
                        matchedPlans.Add(clinicalPlanCandidates[0]);

                        ListViewItem temp_lvi = new ListViewItem();

                        //tumour specific stuff in matches listview (via regex of course name)
                        if (tumourSpecific_checkBox.Checked)
                        {
                            if (brCancer_radioButton.Checked)
                            {
                                int Lat_int = 0;
                                int Loco_int = 0;
                                int Regional_int = 0;

                                string LT_string = "";
                                if (new Regex(".*" + brCaLatLTregex_textBox.Text + ".*", RegexOptions.IgnoreCase).IsMatch(clinicalPlanCandidates[0].Course))
                                {
                                    LT_string = "LT";
                                    Lat_int++;
                                    BrCaLT_int++;
                                }
                                string RT_string = "";
                                if (new Regex(".*" + brCaLatRTregex_textBox.Text + ".*", RegexOptions.IgnoreCase).IsMatch(clinicalPlanCandidates[0].Course))
                                {
                                    RT_string = "RT";
                                    Lat_int++;
                                    BrCaRT_int++;
                                }
                                string BL_string = "";
                                if (new Regex(".*" + brCaLatBLregex_textBox.Text + ".*", RegexOptions.IgnoreCase).IsMatch(clinicalPlanCandidates[0].Course))
                                {
                                    BL_string = "BL";
                                    Lat_int++;
                                    BrCaBL_int++;
                                }
                                
                                string BR_string = "";
                                if (new Regex(".*" + brCaRegionBRregex_textBox.Text + ".*", RegexOptions.IgnoreCase).IsMatch(clinicalPlanCandidates[0].Course))
                                {
                                    BR_string = "BR";
                                    Loco_int++;
                                    BrCaBR_int++;
                                }
                                string CW_string = "";
                                if (new Regex(".*" + brCaRegionCWregex_textBox.Text + ".*", RegexOptions.IgnoreCase).IsMatch(clinicalPlanCandidates[0].Course))
                                {
                                    CW_string = "CW";
                                    Loco_int++;
                                    BrCaCW_int++;
                                }
                                string SCF_string = "";
                                if (new Regex(".*" + brCaRegionSCFregex_textBox.Text + ".*", RegexOptions.IgnoreCase).IsMatch(clinicalPlanCandidates[0].Course))
                                {
                                    SCF_string = "SCF";
                                    Regional_int++;
                                    BrCaSCF_int++;
                                }
                                string AX_string = "";
                                if (new Regex(".*" + brCaRegionAXregex_textBox.Text + ".*", RegexOptions.IgnoreCase).IsMatch(clinicalPlanCandidates[0].Course))
                                {
                                    AX_string = "AX";
                                    Regional_int++;
                                    BrCaAX_int++;
                                }
                                string IMC_string = "";
                                if (new Regex(".*" + brCaRegionIMCregex_textBox.Text + ".*", RegexOptions.IgnoreCase).IsMatch(clinicalPlanCandidates[0].Course))
                                {
                                    IMC_string = "IMC";
                                    Regional_int++;
                                    BrCaIMC_int++;
                                }



                                temp_lvi = new ListViewItem(new string[] { clinicalPlanCandidates[0].MRN, clinicalPlanCandidates[0].Course, clinicalPlanCandidates[0].Plan, clinicalPlanCandidates[0].CourseStart.ToString(), clinicalPlanCandidates[0].Dose.ToString(), clinicalPlanCandidates[0].Fractions.ToString(), LT_string, RT_string, BL_string, Lat_int.ToString(), BR_string, CW_string, Loco_int.ToString(), SCF_string, AX_string, IMC_string, Regional_int.ToString()});
                            }
                        }
                        else
                        {
                           temp_lvi = new ListViewItem(new string[] { clinicalPlanCandidates[0].MRN, clinicalPlanCandidates[0].Course, clinicalPlanCandidates[0].Plan, clinicalPlanCandidates[0].CourseStart.ToString(), clinicalPlanCandidates[0].Dose.ToString(), clinicalPlanCandidates[0].Fractions.ToString() });
                        }


                        temp_lvi.Name = pt.mrn; //give a name to make 'findable'
                        temp_lvi.Tag = clinicalPlanCandidates[0].Structures;
                        Matches_listView.Items.Add(temp_lvi);
                        break;
                    case 0:
                        //ptNode.Nodes.Add("NO APPROPRIATE PLAN - consider loosening filter criteria");
                        noneParentNode.Nodes.Add(ptNode);
                        Console.WriteLine("NONE - NO APPROPRIATE PLAN FOUND");
                        break;
                    default: // >1
                             //ptNode.Nodes.Add("SEVERAL POTENTIAL PLANS REVIEW NEEDED - use override mechanism or tighten filter criteria");
                        reviewParentNode.Nodes.Add(ptNode);
                        Console.WriteLine("REVIEW - SEVERAL POTENTIAL PLANS");
                        break;
                }

            }

            valid_treeView.Nodes.Add(validParentNode);
            review_treeView.Nodes.Add(reviewParentNode);
            none_treeView.Nodes.Add(noneParentNode);
            //archived_treeView.Nodes.Add(archivedParentNode);

            //"bold" path identified structure set in valid tree view
            if (Matches_listView.Items.Count > 0)
            {
                valid_treeView.BeginUpdate();
                foreach (ListViewItem lvi in Matches_listView.Items)
                {
                    string valid_mrn = lvi.SubItems[0].Text;
                    string valid_course_id = lvi.SubItems[1].Text;
                    string valid_plan_id = lvi.SubItems[2].Text;
                    //Bold path to relevant valid structure set from node below patient 
                    BoldThisNodeinThisTreeView(valid_treeView.Nodes[0].Nodes[valid_mrn].Nodes[0], valid_treeView);  //courses node
                    BoldThisNodeinThisTreeView(valid_treeView.Nodes[0].Nodes[valid_mrn].Nodes[0].Nodes[valid_course_id], valid_treeView); //course node
                    BoldThisNodeinThisTreeView(valid_treeView.Nodes[0].Nodes[valid_mrn].Nodes[0].Nodes[valid_course_id].Nodes[0], valid_treeView); //plans node
                    BoldThisNodeinThisTreeView(valid_treeView.Nodes[0].Nodes[valid_mrn].Nodes[0].Nodes[valid_course_id].Nodes[0].Nodes[valid_plan_id], valid_treeView); //plan node
                    BoldThisNodeinThisTreeView(valid_treeView.Nodes[0].Nodes[valid_mrn].Nodes[0].Nodes[valid_course_id].Nodes[0].Nodes[valid_plan_id].Nodes[0], valid_treeView); //structures node
                }
                valid_treeView.EndUpdate();

                // Add sums to tumour specific columns
                if (tumourSpecific_checkBox.Checked)
                {
                    if (brCancer_radioButton.Checked)
                    {
                        Matches_listView.Columns["brCa_LT"].Text = "LT (" + BrCaLT_int.ToString() + ")";
                        Matches_listView.Columns["brCa_RT"].Text = "RT (" + BrCaRT_int.ToString() + ")";
                        Matches_listView.Columns["brCa_BL"].Text = "BL (" + BrCaBL_int.ToString() + ")";
                        Matches_listView.Columns["brCa_BR"].Text = "BR (" + BrCaBR_int.ToString() + ")";
                        Matches_listView.Columns["brCa_CW"].Text = "CW (" + BrCaCW_int.ToString() + ")";
                        Matches_listView.Columns["brCa_SCF"].Text = "SCF (" + BrCaSCF_int.ToString() + ")";
                        Matches_listView.Columns["brCa_AX"].Text = "AX (" + BrCaAX_int.ToString() + ")";
                        Matches_listView.Columns["brCa_IMC"].Text = "IMC (" + BrCaIMC_int.ToString() + ")";
                    }
                }


            }
            Matches_listView.EndUpdate();

            //Expand trees to pt level
            valid_treeView.Nodes[0].Expand();
            review_treeView.Nodes[0].Expand();
            none_treeView.Nodes[0].Expand();
            //archived_treeView.Nodes[0].Expand();

            indicatorValid_label.Text = validParentNode.Nodes.Count.ToString() + " valid";
            indicatorValid_label.Visible = true;
            if (validParentNode.Nodes.Count == 0) { valid_treeView.Nodes.Clear(); }
            indicatorReview_label.Text = reviewParentNode.Nodes.Count.ToString() + " for review";
            indicatorReview_label.Visible = true;
            if (reviewParentNode.Nodes.Count == 0) { review_treeView.Nodes.Clear(); }
            indicatorNone_label.Text = noneParentNode.Nodes.Count.ToString() + " none";
            indicatorNone_label.Visible = true;
            if (noneParentNode.Nodes.Count == 0) { none_treeView.Nodes.Clear(); }
            //indicatorArchived_label.Text = archivedParentNode.Nodes.Count.ToString() + " archived";
            //indicatorArchived_label.Visible = true;
            //if (archivedParentNode.Nodes.Count == 0) { archived_treeView.Nodes.Clear(); }

            Console.WriteLine("Progress: 100.00 %");
            Console.WriteLine("=== FILTER LOCAL DATA (Finish) ===");
        }


        private string exploreDosimetricData()
        {
            int ptMax;
            int cMax;
            int plMax;
            int stMax;
            double prog;
            string interimOutput = "";
            string finalOutput = "";
            TreeNode parentNode;
            TreeNode validParentNode;
            TreeNode reviewParentNode;
            TreeNode noneParentNode;
            TreeNode archivedParentNode;
            TreeNode ptNode;
            TreeNode cNodes;
            TreeNode cNode;
            TreeNode plNodes;
            TreeNode plNode;
            TreeNode infoNode;
            TreeNode stNodes;
            int nodeIndex;

            matchedPlans.Clear();
            valid_treeView.Nodes.Clear();
            review_treeView.Nodes.Clear();
            none_treeView.Nodes.Clear();
            archived_treeView.Nodes.Clear();
            Matches_listView.Items.Clear();
            indicatorValid_label.Visible = false;
            indicatorReview_label.Visible = false;
            indicatorNone_label.Visible = false;
            indicatorArchived_label.Visible = false;

            List<PlanPathAndStructures> clinicalPlanCandidates = new List<PlanPathAndStructures>();
            List<StructureInfo> structureInfoContainer = new List<StructureInfo>();
            bool courseIsMatch;
            bool planIsMatch;
            string excludeCode; //?TODO handle with array to capture multiple exclusion codes per plan
            Regex digit_digit = new Regex(@"\d_\d");

            //ptMax = MRNList.Count;
            ptMax = MRN_listView.Items.Count;

            validParentNode = new TreeNode("VALID - successful filtering");
            reviewParentNode = new TreeNode("REVIEW NEEDED - use override mechanism or tighten filter criteria");
            noneParentNode = new TreeNode("NONE - consider loosening filter criteria");
            archivedParentNode = new TreeNode("ARCHIVED - plan data not available due to archiving");

            for (int ptIterator = 0; ptIterator < ptMax; ptIterator++)
            {
                //interimOutput = "";
                clinicalPlanCandidates.Clear();

                double ptProg = (double)ptIterator / ptMax;
                prog = ptProg;
                Console.WriteLine("Progress: " + prog.ToString("P") + "(Pt #" + (ptIterator + 1).ToString() + " of " + ptMax + ")");

                string MRN = MRN_listView.Items[ptIterator].Text;

                string overriddenCourseID = "";
                string overriddenPlanID = "";
                if (MRN_listView.Items[ptIterator].SubItems.Count > 2) //both course and plan override set
                {
                    overriddenCourseID = MRN_listView.Items[ptIterator].SubItems[1].Text;
                    overriddenPlanID = MRN_listView.Items[ptIterator].SubItems[2].Text;
                }
                bool overridden_flag = false;
                if (overriddenCourseID != "" && overriddenPlanID != "") { overridden_flag = true; }

                if (archivedMRNs.Contains(MRN))
                {
                    //ptNode = parentNode.Nodes.Add("Pt " + (ptIterator + 1).ToString() + " (" + MRN + ")");
                    //ptNode.Nodes.Add("ARCHIVED - plan data not available due to archiving");
                    archivedParentNode.Nodes.Add(MRN, "Pt " + (ptIterator + 1).ToString() + " (" + MRN + ")");
                    Console.WriteLine("ARCHIVED");
                    continue; //skip this patient.
                }

                Patient pt = app.OpenPatientById(MRN);
                //interimOutput += " *** Pt " + (ptIterator + 1).ToString() + " (" + MRN + ") ***" + Environment.NewLine;
                //ptNode = parentNode.Nodes.Add(MRN, "Pt " + (ptIterator + 1).ToString() + " (" + MRN + ")");
                ptNode = new TreeNode("Pt " + (ptIterator + 1).ToString() + " (" + MRN + ")");
                ptNode.Name = MRN;  //put a name so, can later be used for navigation

                cMax = pt.Courses.Count();
                //interimOutput += "  |_ Courses (n=" + cMax.ToString() + ")" + Environment.NewLine;
                cNodes = new TreeNode("Courses (n=" + cMax.ToString() + ")");
                ptNode.Nodes.Add(cNodes);

                for (int cIterator = 0; cIterator < cMax; cIterator++)
                {
                    //reset match flag and exclusionCode
                    courseIsMatch = true;
                    excludeCode = "";

                    double cProg = (double)cIterator / (ptMax * cMax);
                    prog = ptProg + cProg;
                    Console.WriteLine("Progress: " + prog.ToString("P"));

                    Course c = pt.Courses.ToList()[cIterator];
                    if (!c.StartDateTime.Value.IsInRange(From_dateTimePicker.Value, To_dateTimePicker.Value)) { courseIsMatch = false; excludeCode = " => [2a]"; };
                    if (!c.Id.StartsWith(CourseMatchStart_TextBox.Text)) { courseIsMatch = false; excludeCode = " => [3a]"; };
                    if (CourseMatchExcludes_textBox.Lines.Any(s => c.Id.Contains(s))) { courseIsMatch = false; excludeCode = " => [3b]"; };


                    //interimOutput += "    |_ ** " + c.Id + " **" + Environment.NewLine;
                    cNode = cNodes.Nodes.Add(c.Id, c.Id + " (Start date: " + String.Format("{0:d}", c.StartDateTime) + ")" + excludeCode);

                    plMax = c.PlanSetups.Count();
                    //interimOutput += "      |_ Plans (n=" + plMax.ToString() + ")" + Environment.NewLine;
                    plNodes = cNode.Nodes.Add("Plans (n=" + plMax.ToString() + ")");

                    for (int plIterator = 0; plIterator < plMax; plIterator++)
                    {
                        double plProg = (double)plIterator / (ptMax * cMax * plMax);
                        prog = ptProg + cProg + plProg;
                        Console.WriteLine("Progress: " + prog.ToString("P"));

                        //reset match flag and exclusionCode
                        planIsMatch = true;
                        excludeCode = "";

                        PlanSetup pl = c.PlanSetups.ToList()[plIterator];
                        if (replanIdentification_checkBox.Checked)
                        {
                            if (digit_digit.IsMatch(pl.Id)) { planIsMatch = false; excludeCode = " => [4a]"; };
                        }

                        if (advFilters_checkBox.Checked && planRegex_textBox.Text != "")
                        {
                            Regex planRegex = new Regex(@planRegex_textBox.Text);
                            if (planRegex.IsMatch(pl.Id)) { planIsMatch = false; excludeCode = " => [4b]"; };
                        }


                        //interimOutput += "        |_ * " + pl.Id + " *" + Environment.NewLine;
                        plNode = plNodes.Nodes.Add(pl.Id, pl.Id + excludeCode); //+ " (Plan Type: " + pl.PlanType + "; Pl Approver: " + pl.PlanningApprover + "; Tx Approver: " + pl.TreatmentApprover + ")");


                        //reset excludeCode
                        excludeCode = "";

                        //bool dose_check_flag = true;  //whether to take isDoseValid into account (can rule out some old plans WITH STRUCTURES AND DOSE ?old dose calc algorithm) 
                        if (validDose_checkBox.Checked)
                        {
                            if (!pl.IsDoseValid)
                            {
                                //dose_check_flag = false;
                                planIsMatch = false;
                                excludeCode = " => [5a]";
                            }
                        }

                        if (!pl.IsTreated) //might need similar approach via is_treated_flag than isDoseValid above
                        {
                            planIsMatch = false;
                            excludeCode = " => [5b]";
                        }

                        if (pl.StructureSet != null) //needed to avoid structureless plans when relaxing isDoseValid and is isTreated conditions
                        {
                            StructureSet ss = pl.StructureSet;
                            stMax = ss.Structures.Count();
                            //interimOutput += "          |_ Structures (n=" + stMax.ToString() + ")" + Environment.NewLine;
                            stNodes = plNode.Nodes.Add("Structures (n=" + stMax.ToString() + ")");

                            if (stMax > 0)
                            {
                                structureInfoContainer.Clear();
                                for (int stIterator = 0; stIterator < stMax; stIterator++)
                                {

                                    Structure st = ss.Structures.ToList()[stIterator];
                                    //interimOutput += "            |_ * " + st.Id + " *" + Environment.NewLine;
                                    if (st.DicomType != "MARKER")  //could potentially lead to no added structures despite structure count > 0 
                                    {
                                        structureInfoContainer.Add(new StructureInfo { Id = st.Id, isEmpty = st.IsEmpty });
                                        stNodes.Nodes.Add(st.Id);
                                    }
                                }
                                if (structureInfoContainer.Count() == 0) { planIsMatch = false; excludeCode = " => [6a]"; } //accounting for case where all structures are MARKERs

                            }
                            else { planIsMatch = false; excludeCode = " => [6a]"; }

                        }
                        else { planIsMatch = false; excludeCode = " => [6a]"; }

                        if (overridden_flag)
                        {

                            if (overriddenCourseID == c.Id && overriddenPlanID == pl.Id)  //if overridden values are valid always add as candidats
                            {
                                if (structureInfoContainer.Count() > 0)
                                {

                                    clinicalPlanCandidates.Add(new PlanPathAndStructures { MRN = MRN, Course = c.Id, CourseStart = c.StartDateTime, Plan = pl.Id, Structures = new List<StructureInfo>(structureInfoContainer) });
                                    courseIsMatch = false; //set to make sure doesn't get added twice in next block
                                }
                                else
                                {
                                    MessageBox.Show("Overriden plan with MRN '" + MRN + "' does not have structures attached.", "No attached structures", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1);
                                }
                            }

                        }

                        if (courseIsMatch && planIsMatch)  // if matches filter criteria
                        {
                            clinicalPlanCandidates.Add(new PlanPathAndStructures { MRN = MRN, Course = c.Id, CourseStart = c.StartDateTime, Plan = pl.Id, Structures = new List<StructureInfo>(structureInfoContainer) });
                        }


                        infoNode = plNode.Nodes.Add("Plan Meta Info" + excludeCode);
                        infoNode.Nodes.Add("Type: " + pl.PlanType);
                        infoNode.Nodes.Add("Intent: " + pl.PlanIntent);
                        infoNode.Nodes.Add("Planning Approver: " + pl.PlanningApprover);
                        infoNode.Nodes.Add("Tx Approver: " + pl.TreatmentApprover);
                        infoNode.Nodes.Add("ValidDose Flag: " + pl.IsDoseValid.ToString());
                        infoNode.Nodes.Add("Treated Flag: " + pl.IsTreated.ToString());
                        if (courseIsMatch && planIsMatch)  //some physics plans caused crash, if ongoing issue consider to remove check for this type of meta data
                        {
                            if (pl.VerifiedPlan != null)
                            {
                                infoNode.Nodes.Add("Verified Plan (?self): " + pl.VerifiedPlan.Id);
                            }
                            else
                            {
                                infoNode.Nodes.Add("Verified Plan (?self): <null>");
                            }
                        }
                        else
                        {
                            infoNode.Nodes.Add("Verified Plan (?self): not checked");
                        }

                        if (pl.StructureSet != null)
                        {
                            infoNode.Nodes.Add("Attached structures (count): " + pl.StructureSet.Structures.Count().ToString());
                        }
                        else
                        {
                            infoNode.Nodes.Add("Attached structures (count): <null>");
                        }

                    }
                }


                switch (clinicalPlanCandidates.Count)
                {
                    case 1:
                        //ptNode.Nodes.Add("SINGLE APPROPRIATE PLAN - successful filtering");
                        validParentNode.Nodes.Add(ptNode);
                        Console.WriteLine("VALID - SINGLE APPROPRIATE PLAN");
                        matchedPlans.Add(clinicalPlanCandidates[0]);

                        ListViewItem temp_lvi = new ListViewItem(new string[] { clinicalPlanCandidates[0].MRN, clinicalPlanCandidates[0].Course, clinicalPlanCandidates[0].Plan });
                        temp_lvi.Name = MRN; //give a name to make 'findable'
                        temp_lvi.Tag = clinicalPlanCandidates[0].Structures;
                        Matches_listView.Items.Add(temp_lvi);
                        break;
                    case 0:
                        //ptNode.Nodes.Add("NO APPROPRIATE PLAN - consider loosening filter criteria");
                        noneParentNode.Nodes.Add(ptNode);
                        Console.WriteLine("NONE - NO APPROPRIATE PLAN FOUND");
                        break;
                    default: // >1
                             //ptNode.Nodes.Add("SEVERAL POTENTIAL PLANS REVIEW NEEDED - use override mechanism or tighten filter criteria");
                        reviewParentNode.Nodes.Add(ptNode);
                        Console.WriteLine("REVIEW - SEVERAL POTENTIAL PLANS");
                        break;
                }

                finalOutput += interimOutput;
                app.ClosePatient();
            }




            valid_treeView.Nodes.Add(validParentNode);
            review_treeView.Nodes.Add(reviewParentNode);
            none_treeView.Nodes.Add(noneParentNode);
            archived_treeView.Nodes.Add(archivedParentNode);

            //"bold" path identified structure set in valid tree view
            if (Matches_listView.Items.Count > 0)
            {
                valid_treeView.BeginUpdate();
                foreach (ListViewItem lvi in Matches_listView.Items)
                {
                    string valid_mrn = lvi.SubItems[0].Text;
                    string valid_course_id = lvi.SubItems[1].Text;
                    string valid_plan_id = lvi.SubItems[2].Text;
                    //Bold path to relevant valid structure set from node below patient 
                    BoldThisNodeinThisTreeView(valid_treeView.Nodes[0].Nodes[valid_mrn].Nodes[0], valid_treeView);  //courses node
                    BoldThisNodeinThisTreeView(valid_treeView.Nodes[0].Nodes[valid_mrn].Nodes[0].Nodes[valid_course_id], valid_treeView); //course node
                    BoldThisNodeinThisTreeView(valid_treeView.Nodes[0].Nodes[valid_mrn].Nodes[0].Nodes[valid_course_id].Nodes[0], valid_treeView); //plans node
                    BoldThisNodeinThisTreeView(valid_treeView.Nodes[0].Nodes[valid_mrn].Nodes[0].Nodes[valid_course_id].Nodes[0].Nodes[valid_plan_id], valid_treeView); //plan node
                    BoldThisNodeinThisTreeView(valid_treeView.Nodes[0].Nodes[valid_mrn].Nodes[0].Nodes[valid_course_id].Nodes[0].Nodes[valid_plan_id].Nodes[0], valid_treeView); //structures node
                }
                valid_treeView.EndUpdate();
            }
            //Expand trees to pt level
            valid_treeView.Nodes[0].Expand();
            review_treeView.Nodes[0].Expand();
            none_treeView.Nodes[0].Expand();
            archived_treeView.Nodes[0].Expand();


            indicatorValid_label.Text = validParentNode.Nodes.Count.ToString() + " valid";
            indicatorValid_label.Visible = true;
            if (validParentNode.Nodes.Count == 0) { valid_treeView.Nodes.Clear(); }
            indicatorReview_label.Text = reviewParentNode.Nodes.Count.ToString() + " for review";
            indicatorReview_label.Visible = true;
            if (reviewParentNode.Nodes.Count == 0) { review_treeView.Nodes.Clear(); }
            indicatorNone_label.Text = noneParentNode.Nodes.Count.ToString() + " none";
            indicatorNone_label.Visible = true;
            if (noneParentNode.Nodes.Count == 0) { none_treeView.Nodes.Clear(); }
            indicatorArchived_label.Text = archivedParentNode.Nodes.Count.ToString() + " archived";
            indicatorArchived_label.Visible = true;
            if (archivedParentNode.Nodes.Count == 0) { archived_treeView.Nodes.Clear(); }


            Console.WriteLine("Progress: 100.00 %");
            return (finalOutput);
        }

        private void DVHExport_Button_Click(object sender, EventArgs e)
        {
            //this.OutputFile_TextBox.Text = "";
            if (MRN_listView.Items.Count > 0)
            {
                OutputFile_TextBox.Text = exploreDosimetricData();
            }

            //string OutputDir = @"C:\Temp\DVHExporter";
            //System.IO.Directory.CreateDirectory(OutputDir);
            //OutputFile_TextBox.Text = "";

            //foreach (string MRN in MRNList)
            //{
            //    Patient p = app.OpenPatientById(MRN);
            //    foreach (Course c in p.Courses)
            //    {

            //        // select plans from the course that have valid dose
            //        foreach (PlanSetup plan in c.PlanSetups)
            //        {
            //            if (plan.IsDoseValid && plan.IsTreated)
            //            {
            //                //Find all structures lniked to this plan:
            //                StructureSet structureset = plan.StructureSet;
            //                if (structureset.Structures.Count() > 0)
            //                {
            //                    foreach (Structure structure in structureset.Structures)
            //                    {
            //                        if (structure != null)
            //                        {
            //                            try
            //                            {
            //                                // extract the DVH of the structure and dump to a file:
            //                                DVHData DVH = plan.GetDVHCumulativeData(structure, DoseValuePresentation.Relative, VolumePresentation.Relative, 0.1);
            //                                string FileName = string.Format(@"{0}\MRN-{1}_Course-{2}_Plan-{3}_StructSet-{4}_Struct-{5}.txt", OutputDir, p.Id, c.Id, plan.Id, structureset.Id, structure.Id);
            //                                //WriteDVHFile(UserID, FileName, MRN, p, c, plan, structure, DVH);



            //                                //OutputFile_ListBox.Items.Add(FileName);
            //                            }
            //                            catch (Exception exception)
            //                            {
            //                                //MessageBox.Show("Exception was thrown:" + exception.Message);
            //                            }

            //                        }

            //                    }

            //                    // Save this to list of output files in GUI:
            //                    OutputFile_ListBox.Items.Add(string.Format(@"MRN-{0}_Course-{1}_Plan-{2}_StructSet-{3}", p.Id, c.Id, plan.Id, structureset.Id));
            //                    OutputFile_TextBox.Text += string.Format(@"MRN-{0}_Course-{1}_Plan-{2}_StructSet-{3}", p.Id, c.Id, plan.Id, structureset.Id) + Environment.NewLine;
            //                    //Items.Add(string.Format(@"MRN-{0}_Course-{1}_Plan-{2}_StructSet-{3}", p.Id, c.Id, plan.Id, structureset.Id));
            //                }
            //            }
            //        }
            //    }
            //    app.ClosePatient();
            //}


            // The following is just an example; export for all patients:
            //var PatSummaries = app.PatientSummaries;
            //foreach (PatientSummary ps in PatSummaries)
            //{
            //    Patient p = app.OpenPatient(ps);
            //    foreach (Course c in p.Courses)
            //    {
            //        // select plans from the course that have valid dose
            //        foreach (PlanSetup plan in c.PlanSetups.Where(x => x.IsDoseValid))
            //        {
            //            // find the planning target
            //            Structure target = plan.StructureSet.Structures.Where(x => x.Id == plan.TargetVolumeID).FirstOrDefault();
            //            // extract the DVH of the planning target, dump to a file in c:\temp\dvhdump directory.
            //            if (plan.Dose != null && target != null)
            //            {
            //                DVHData dvh = plan.GetDVHCumulativeData(target, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
            //                string filename = string.Format(@"{0}\{1}_{2}_{3}_{4}-dvh.csv",
            //                    OutputDir, p.Id, c.Id, plan.Id, target.Id);
            //                DumpDVH(filename, dvh);
            //                Console.WriteLine(filename);
            //            }
            //        }
            //    }
            //    app.ClosePatient();
            //}

        }

        private void Login_Button_Click(object sender, EventArgs e)
        {

        }

        private void SelectMRNList_Button_Click(object sender, EventArgs e)
        {
            // Clear list boxes:
            this.MRN_ListBox.Items.Clear();
            MRN_listView.Items.Clear();
            mrnList_label.Text = "MRN List";
            this.OutputFile_TextBox.Text = "";
            MRNList.Clear();


            // Load MRNList
            System.Windows.Forms.OpenFileDialog OpenFileDialog = new System.Windows.Forms.OpenFileDialog();
            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\mrns"))
            {
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\data\mrns");
            }
            OpenFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\mrns";
            OpenFileDialog.Filter = "TEXT files (*.txt)|*.txt|All files (*.*)|*.*";
            try
            {

                if (OpenFileDialog.ShowDialog() == DialogResult.OK)
                {

                    bool AnyDigitsAppended = false;
                    bool AnyDigitsTruncated = false;
                    string FileName = OpenFileDialog.FileName;
                    string[] FileDataString = File.ReadAllLines(FileName);
                    int N = FileDataString.Count();
                    for (int i = 0; i < N; i++)
                    {
                        var parts = FileDataString[i].Split('\t');

                        string MRNtemp = parts[0];

                        while (MRNtemp.Length < 7)
                        {
                            MRNtemp = MRNtemp.Insert(0, "0");
                            AnyDigitsAppended = true;
                        }
                        while (MRNtemp.Length > 7)
                        {
                            MRNtemp = MRNtemp.Remove(0, 1);
                            AnyDigitsTruncated = true;
                        }

                        MRNList.Add(MRNtemp);
                        ListViewItem lvi = new ListViewItem(MRNtemp); //sets Text property
                        lvi.Name = MRNtemp; //sets Name property (to allow searching via Find)
                        if (parts.Count() == 3 || parts.Count() > 3)
                        {
                            lvi.SubItems.Add(parts[1]);
                            lvi.SubItems.Add(parts[2]);
                        }
                        else
                        {
                            lvi.SubItems.Add(" ");
                            lvi.SubItems.Add(" ");
                        }

                        MRN_listView.Items.Add(lvi);
                        MRN_ListBox.Items.Add(MRNtemp);
                    }
                    if (AnyDigitsAppended)
                        MessageBox.Show("Note: At least one MRN had less than 7-digits.");
                    if (AnyDigitsTruncated)
                        MessageBox.Show("Note: At least one MRN had more than 7-digits.");

                    if (MRNList.Count > 0)
                    {
                        mrnList_label.Text = "MRN List (n = " + MRNList.Count.ToString() + ")";
                        //EnableButton(DVHExport_Button);
                    }
                    if (MRNList.Count == 0)
                    {
                        MessageBox.Show("No MRNs in file!");
                        mrnList_label.Text = "MRN List";
                        //DisableButton(DVHExport_Button);
                    }
                }
            }
            catch (Exception errorMsg)
            {
                MessageBox.Show(errorMsg.Message, "Error reading a file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        static void WriteDVHFile(string UserID, string FileName, string MRN, Patient p, Course c, PlanSetup plan, Structure structure, DVHData DVH)
        {
            System.IO.StreamWriter DVHFile = new System.IO.StreamWriter(FileName);
            // write a header
            DVHFile.WriteLine("Patient Name         : " + p.LastName + ", " + p.FirstName);
            DVHFile.WriteLine("Patient ID           : " + p.Id);
            DVHFile.WriteLine("Comment              : Single structure DVH");
            DVHFile.WriteLine("Date                 : " + DateTime.Now.ToString("yyyy/MM/dd"));
            DVHFile.WriteLine("Exported by          : " + UserID);
            DVHFile.WriteLine("Type                 : Cumulative Dose Volume Histogram");
            DVHFile.WriteLine("Description          : The cumulative DVH displays the percentage (relative)");
            DVHFile.WriteLine("                       or volume (absolute) of structures that receive a dose");
            DVHFile.WriteLine("                       equal to or greater than a given dose.");
            DVHFile.WriteLine("");
            DVHFile.WriteLine("Plan: " + plan.Id);
            DVHFile.WriteLine("Course: " + c.Id);
            if (plan.IsTreated)
            { DVHFile.WriteLine("Plan Status: Completed"); }
            else
            { DVHFile.WriteLine("Plan Status: ???"); }   // < -- ??
            DVHFile.WriteLine("Prescribed dose [cGy]: " + Math.Round(plan.TotalPrescribedDose.Dose, 1));
            DVHFile.WriteLine("% for dose(%): ??? "); // < -- ??
            DVHFile.WriteLine("");
            DVHFile.WriteLine("Structure: " + structure.Id);
            DVHFile.WriteLine("Approval Status: " + plan.ApprovalStatus);
            DVHFile.WriteLine("Plan: " + plan.Id);
            DVHFile.WriteLine("Course: " + c.Id);
            DVHFile.WriteLine("Volume [cm³]: " + Math.Round(DVH.Volume, 1));
            DVHFile.WriteLine("Dose Cover.[%]: " + Math.Round(100 * DVH.Coverage, 1));
            DVHFile.WriteLine("Sampling Cover.[%]: " + Math.Round(100 * DVH.SamplingCoverage, 1));
            DVHFile.WriteLine("Min Dose [%]: " + Math.Round(DVH.MinDose.Dose, 1));
            DVHFile.WriteLine("Max Dose [%]: " + Math.Round(DVH.MaxDose.Dose, 1));
            DVHFile.WriteLine("Mean Dose [%]: " + Math.Round(DVH.MeanDose.Dose, 1));
            DVHFile.WriteLine("Modal Dose [%]: ???"); // < -- ??
            DVHFile.WriteLine("Median Dose [%]: " + Math.Round(DVH.MedianDose.Dose, 1));
            DVHFile.WriteLine("STD [%]: " + Math.Round(DVH.StdDev, 1));
            DVHFile.WriteLine("Equiv. Sphere Diam. [cm]: " + Math.Round(2 * Math.Pow((3 * DVH.Volume / 4 / Math.PI), (Convert.ToDouble(1) / Convert.ToDouble(3))), 1));
            DVHFile.WriteLine("Conformity Index: ???"); // < -- ??
            DVHFile.WriteLine("Gradient Measure [cm]: ???"); // < -- ??
            DVHFile.WriteLine("");
            DVHFile.WriteLine("Relative dose [%]          Dose [cGy] Ratio of Total Structure Volume [%]");
            DVHFile.WriteLine("");

            // write all dvh points for the PTV.
            foreach (DVHPoint pt in DVH.CurveData)
            {
                string line = string.Format("{0} {1} {2}", Math.Round(pt.DoseValue.Dose, 1), Math.Round(plan.TotalPrescribedDose.Dose * pt.DoseValue.Dose / 100, 0), Convert.ToSingle(pt.Volume));
                DVHFile.WriteLine(line);
            }
            DVHFile.Close();
        }

        private void EnableButton(Button button)
        {
            button.Enabled = true;
            button.BackColor = Color.WhiteSmoke;
            button.ForeColor = SystemColors.ActiveCaptionText;
        }

        private void DisableButton(Button button)
        {
            button.Enabled = false;
            button.BackColor = Color.LightGray;
            button.ForeColor = SystemColors.InactiveCaptionText;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void CourseMatchExcludes_textBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void replanIdentification_checkBox_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void CourseMatchStart_TextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void Analyse_button_Click(object sender, EventArgs e)
        {
            //Results_tabPage.Text = "Results - across cohort ( n=" + matchedPlans.Count.ToString() + " )";
            Results_listView.Items.Clear();
            Results_listView.Groups.Clear();
            if (groupByDate_checkBox.Checked)
            {
                for (int i = 0; i < dateRanges_listView.Items.Count; i++)
                {
                    ListViewItem lvi = dateRanges_listView.Items[i];
                    if (perStrucV2_radioButton.Checked)
                    { AlternativeAnalyseNames2(true, lvi.Text, Convert.ToDateTime(lvi.SubItems[1].Text), Convert.ToDateTime(lvi.SubItems[2].Text), Results_listView); }
                    else
                    { AlternativeAnalyseNames(true, lvi.Text, Convert.ToDateTime(lvi.SubItems[1].Text), Convert.ToDateTime(lvi.SubItems[2].Text), Results_listView); }
                }
                if (perStrucV2_radioButton.Checked)
                { AlternativeAnalyseNames2(false, "All", null, null, Results_listView); }
                else
                { AlternativeAnalyseNames(false, "All", null, null, Results_listView); }
            }
            else
            {
                if (perStrucV2_radioButton.Checked)
                { AlternativeAnalyseNames2(false, "", null, null, Results_listView); }
                else
                { AlternativeAnalyseNames(false, "", null, null, Results_listView); }     
            }
            //AnalyseNames(true, "sept 2012", new DateTime (2012, 09, 01), new DateTime(2012, 09, 30), Results_listView);
            //AnalyseNames(true, "oct-dec 2012", new DateTime(2012, 10, 01), new DateTime(2012, 12, 31), Results_listView);
            //AnalyseNames(true, "jan-aug 2012", new DateTime(2012, 01, 01), new DateTime(2012, 08, 31), Results_listView);
        }

        private void AlternativeAnalyseNames(bool perDateRange, string dateRangeLabel, DateTime? rangeStart, DateTime? rangeEnd, ListView resultsLV)
        {

            if (matchedPlans.Count > 0 && Template_listView.Items.Count > 0)
            {
                //Results_textBox.Text = "";
                //Results_textBox.Text += "TEST" + Environment.NewLine;
                //foreach (this.MRN_listView.Items.)

                resultsLV.BeginUpdate();
                //ListViewGroup default_lvg = new ListViewGroup("Default (no group set)", HorizontalAlignment.Left);


                foreach (ListViewItem lvi in Template_listView.Items)  // iterate through structures in template
                {
                    string tmplt_structure_name = lvi.SubItems[0].Text;
                    string tmplt_structure_group = "";
                    if (lvi.SubItems.Count > 1)
                    {
                        tmplt_structure_group = lvi.SubItems[1].Text;
                    }

                    //Create set of synonyms
                    HashSet<string> tmplt_synonymes = new HashSet<string>(); //String.Comparer.OrdinalIgnoreCase
                    if (lvi.SubItems.Count > 2)
                    {
                        List<string> temp_synonymes = new List<string>();
                        temp_synonymes = lvi.SubItems[2].Text.Split(',').ToList();
                        foreach (string temp_synonym in temp_synonymes)
                        {
                            tmplt_synonymes.Add(temp_synonym.Trim());
                        }
                    }

                    int adherance_count = 0;
                    int syn_adherance_count = 0;
                    int empty_count = 0;
                    int syn_empty_count = 0;
                    int missing_count = 0;
                    int range_count = 0; // number of plans in specified date range
                    int auto_map_count = 0;
                    int auto_exact_count = 0;
                    int auto_synonym_count = 0;
                    int manual_map_count = 0;
                    int no_map_count = 0;

                    foreach (PlanPathAndStructures ppas in matchedPlans) // compare to valid plans
                    {
                        if (perDateRange)
                        {

                            if (!ppas.CourseStart.Value.IsInRange((DateTime)rangeStart, (DateTime)rangeEnd))
                            {
                                continue; // skip to next ppas
                            }
                            else { range_count = range_count + 1; }
                        }
                        else
                        {
                            range_count = matchedPlans.Count;
                        }

                        bool adherance_flag = false; // TRUE IF: template structure name match AND this contour is NOT empty
                        bool empty_flag = false; // TRUE IF: template structure name match AND this contour is empty; should be mutually exclusive with adherence_flag
                        int syn_adherance_flag = 0; //number of non-empty synonumes structures in a plan; "TRUE" if > 0
                        int syn_empty_flag = 0; // number of empty synonumes structures in a plan; "TRUE" if > 0

                        //foreach (StructureInfo si in ppas.Structures)  // iterate over all structures of each valid plans
                        //{

                        //    if (si.Id.Trim() == tmplt_structure_name.Trim()) //check for structure name match
                        //    {
                        //        if (si.isEmpty)
                        //        {
                        //            empty_count = empty_count + 1;
                        //            empty_flag = true;
                        //        }
                        //        else
                        //        {
                        //            adherance_count = adherance_count + 1;
                        //            adherance_flag = true;
                        //        }
                        //    }

                        //    else  //if no match check whether synonyms match
                        //    {
                        //        for (int i = 0; i < synonymes.Count; i++)  //iterate over synonyms of current template row
                        //        {
                        //            if (si.Id == synonymes[i].Trim())
                        //            {
                        //                if (si.isEmpty)
                        //                {
                        //                    syn_empty_flag = true;
                        //                    syn_empty_count = syn_empty_count + 1;
                        //                }
                        //                else
                        //                {
                        //                    syn_adherance_flag = true;
                        //                    syn_adherance_count = syn_adherance_count + 1;
                        //                }
                        //                break; //if synonym match found no need to look further in synonyms list 
                        //            }
                        //        }
                        //    }

                        //    if (adherance_flag | empty_flag | syn_adherance_flag | syn_empty_flag)
                        //    {
                        //        break; //if structure_name match of some sort (wheter synonym or empty) found no need to look further in same plan
                        //    }
                        //}


                        HashSet<string> pln_non_empty_strucs = new HashSet<string>();
                        HashSet<string> pln_empty_strucs = new HashSet<string>();
                        foreach (StructureInfo si in ppas.Structures)  // iterate over all structures of each valid plans
                        {

                            if (si.isEmpty)
                            {
                                pln_empty_strucs.Add(si.Id.Trim());
                            }
                            else
                            {
                                pln_non_empty_strucs.Add(si.Id.Trim());
                            }

                            //if (adherance_flag | empty_flag | syn_adherance_flag | syn_empty_flag)
                            //{
                            //    break; //if structure_name match of some sort (wheter synonym or empty) found no need to look further in same plan
                            //}
                        }

                        int not_missing_flag = 0;  //use an int as flag for debug purposes see messagebox below
                        //if (pln_non_empty_strucs.Contains(tmplt_structure_name, StringComparer.OrdinalIgnoreCase))
                        if (pln_non_empty_strucs.Contains(tmplt_structure_name))
                        {
                            adherance_flag = true;
                            adherance_count++;
                            auto_exact_count++;  //
                            not_missing_flag++;
                        }
                        //if (pln_empty_strucs.Contains(tmplt_structure_name, StringComparer.OrdinalIgnoreCase))
                        if (pln_empty_strucs.Contains(tmplt_structure_name))
                        {
                            empty_flag = true;
                            empty_count++;
                            not_missing_flag++;
                        }
                        if (not_missing_flag > 1) { MessageBox.Show("Structure can't be empty and non-empty"); }; //Should never happen

                        HashSet<string> tmp_intersection = new HashSet<string>();
                        tmp_intersection.UnionWith(pln_non_empty_strucs); //temp fill with plan non empty strucs
                        tmp_intersection.IntersectWith(tmplt_synonymes); //remove all but intersecting
                        syn_adherance_flag = tmp_intersection.Count;
                        if (syn_adherance_flag > 0) { syn_adherance_count++; }

                        if (pln_empty_strucs.Count > 0)
                        {
                            tmp_intersection.Clear();
                            tmp_intersection.UnionWith(pln_empty_strucs); //temp fill with plan empty strucs
                            tmp_intersection.IntersectWith(tmplt_synonymes); //remove all but intersecting
                            syn_empty_flag = tmp_intersection.Count;
                        }
                        if (syn_empty_flag > 0) { syn_empty_count++; }

                        if (!adherance_flag & !empty_flag & syn_adherance_flag == 0 & syn_empty_flag == 0) { missing_count = missing_count + 1; };

                        if (adherance_flag | ((!adherance_flag) & syn_adherance_flag == 1)) { auto_map_count++; }
                        if (!adherance_flag & syn_adherance_flag == 1) { auto_synonym_count++; }
                        if (!adherance_flag & syn_adherance_flag == 0) { no_map_count++; }
                    }

                    manual_map_count = range_count - auto_map_count - no_map_count;

                    double average_adherance = (double)adherance_count / range_count;
                    double average_empty = (double)empty_count / range_count;
                    double average_syn_adherance = (double)syn_adherance_count / range_count;
                    double average_syn_empty = (double)syn_empty_count / range_count;
                    double average_missing = (double)missing_count / range_count;
                    double average_auto_map = (double)auto_map_count / range_count;
                    double average_auto_exact = (double)auto_exact_count / range_count;
                    double average_auto_synonym = (double)auto_synonym_count / range_count;
                    double average_manual_map = (double)manual_map_count / range_count;
                    double average_no_map = (double)no_map_count / range_count;

                    ListViewGroup results_lvg = null;
                    if (dateRangeLabel != "")  // could be simplified as NOW groups are known before and (via dateRanges_listView)
                    {
                        bool group_exists = false;
                        foreach (ListViewGroup lvg in resultsLV.Groups)
                        {
                            if (lvg.Header == dateRangeLabel + " [n=" + range_count + "]")
                            {
                                group_exists = true;
                                results_lvg = lvg;
                                break;
                            }
                        }
                        if (!group_exists)
                        {
                            int results_lvg_index = resultsLV.Groups.Add(new ListViewGroup(dateRangeLabel + " [n=" + range_count + "]", HorizontalAlignment.Left));
                            results_lvg = resultsLV.Groups[results_lvg_index];
                        }
                    }



                    ListViewItem results_lvi = new ListViewItem(new string[] { dateRangeLabel, tmplt_structure_name, tmplt_structure_group, adherance_count.ToString(), average_adherance.ToString("P"), empty_count.ToString(), average_empty.ToString("P"), syn_adherance_count.ToString(), average_syn_adherance.ToString("P"), syn_empty_count.ToString(), average_syn_empty.ToString("P"), missing_count.ToString(), average_missing.ToString("P"), auto_map_count.ToString(), average_auto_map.ToString("P"), auto_exact_count.ToString(), average_auto_exact.ToString("P"), auto_synonym_count.ToString(), average_auto_synonym.ToString("P"), manual_map_count.ToString(), average_manual_map.ToString("P"), no_map_count.ToString(), average_no_map.ToString("P") });
                    if (dateRangeLabel != "") { results_lvi.Group = results_lvg; }
                    resultsLV.Items.Add(results_lvi);
                }
                //if (dateRangeLabel != "") { resultsLV.Groups.Add(default_lvg); } //add at end so 'default' group appears last

                resultsLV.EndUpdate();
            }
        }

        private void AlternativeAnalyseNames2(bool perDateRange, string dateRangeLabel, DateTime? rangeStart, DateTime? rangeEnd, ListView resultsLV)
        {
            if (matchedPlans.Count > 0 && RepTemp_objectListView.Items.Count > 0)
            {
                resultsLV.BeginUpdate();

                foreach (ListViewItem lvi in RepTemp_objectListView.Items)  // iterate through structures in template
                {
                    string tmplt_structure_name = lvi.SubItems[0].Text;
                    string tmplt_structure_group = "";
                    if (lvi.SubItems.Count > 1)
                    {
                        tmplt_structure_group = lvi.SubItems[1].Text;
                    }

                    //Create dictonary of synonyms +/- dose/vol rules
                    Dictionary<string, string> synonyms_with_details = new Dictionary<string, string>();
                    //HashSet<string> tmplt_synonyms = new HashSet<string>(); //String.Comparer.OrdinalIgnoreCase
                    if (lvi.SubItems.Count > 2)
                    {
                        List<string> temp_synonyms = new List<string>();
                        temp_synonyms = lvi.SubItems[2].Text.Split(',').ToList();
                        foreach (string syn_details in temp_synonyms)
                        {
                            if (syn_details.Contains("{"))
                            {
                                string temp_syn = "";
                                string temp_details = "";
                                temp_syn = syn_details.Split('{')[0].Trim();
                                temp_details = syn_details.Split('{')[1].Trim(new char[] { '}', ' ' });
                                synonyms_with_details.Add(temp_syn, temp_details);
                            }
                            else
                            {
                                synonyms_with_details.Add(syn_details, null);
                            }
                        }
                    }

                    int adherance_count = 0;
                    int syn_adherance_count = 0;
                    int empty_count = 0;
                    int syn_empty_count = 0;
                    int missing_count = 0;
                    int range_count = 0; // number of plans in specified date range
                    int auto_map_count = 0;
                    int auto_exact_count = 0;
                    int auto_synonym_count = 0;
                    int manual_map_count = 0;
                    int no_map_count = 0;

                    foreach (PlanPathAndStructures ppas in matchedPlans) // compare to valid plans
                    {
                        if (perDateRange)
                        {

                            if (!ppas.CourseStart.Value.IsInRange((DateTime)rangeStart, (DateTime)rangeEnd))
                            {
                                continue; // skip to next ppas
                            }
                            else { range_count = range_count + 1; }
                        }
                        else
                        {
                            range_count = matchedPlans.Count;
                        }

                        bool adherance_flag = false; // TRUE IF: template structure name match AND this contour is NOT empty
                        bool empty_flag = false; // TRUE IF: template structure name match AND this contour is empty; should be mutually exclusive with adherence_flag
                        int syn_adherance_flag = 0; //number of non-empty synonumes structures in a plan; "TRUE" if > 0
                        int syn_empty_flag = 0; // number of empty synonumes structures in a plan; "TRUE" if > 0

 
                        foreach (StructureInfo si in ppas.Structures)  // iterate over all structures of each valid plans
                        {
                            int not_missing_flag = 0;  //use an int as flag for debug purposes see messagebox below

                            if (si.isEmpty)
                            {
                                if (si.Id.Trim() == tmplt_structure_name.Trim())
                                {
                                    empty_flag = true;
                                    empty_count++;
                                    not_missing_flag++;
                                }
                            }
                            else
                            {
                                if (si.Id.Trim() == tmplt_structure_name.Trim())
                                {
                                    adherance_flag = true;
                                    adherance_count++;
                                    auto_exact_count++; 
                                    not_missing_flag++;
                                }
                            }

                            if (not_missing_flag > 1) { MessageBox.Show("Structure can't be empty and non-empty"); }; //Should never happen

                            //foreach (string tmplt_syn in tmplt_synonyms)
                            //{
                            //    if (si.isEmpty)
                            //    {
                            //        if (si.Id.Trim() == tmplt_syn.Trim())
                            //        {
                            //            syn_empty_flag++;  //individual patient count
                            //        }
                            //    }
                            //    else
                            //    {
                            //        if (si.Id.Trim() == tmplt_syn.Trim())
                            //        {
                            //            syn_adherance_flag++;  //individual patient count
                            //        }
                            //    }
                            //}

                            //synonyms
                            for (int i = 0; i < synonyms_with_details.Count(); i++)
                            {
                                if (synonyms_with_details.Keys.ElementAt(i).Trim() == si.Id)
                                {
                                    if (synonyms_with_details.Values.ElementAt(i) != null)
                                    {
                                        bool isConditionMet = false;
                                        var jintEval = new Engine()  //check details expression via JINT (needs to evaluate to true or false)
                                        .SetValue("Vol", si.vol)
                                        .SetValue("Dmean", si.dMean)
                                        .SetValue("D95", si.d95)
                                        .Execute(synonyms_with_details.Values.ElementAt(i)) // details expression of current synonym
                                        .GetCompletionValue()
                                        .ToObject();
                                        isConditionMet = (bool)jintEval;

                                        if (isConditionMet)
                                        {
                                            if (si.isEmpty)
                                            {
                                                syn_empty_flag++;
                                            }
                                            else
                                            {
                                                syn_adherance_flag++;
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine("Matrix - MRN '" + ppas.MRN + "', Structure '" + si.Id + "' (Vol: "
                                          + si.vol.ToString() + "; Dmean: "
                                          + si.dMean.ToString() + ")  does not meet: '"
                                          + synonyms_with_details.Values.ElementAt(i) + "'");
                                        }
                                    }
                                    else
                                    {
                                        if (si.isEmpty)
                                        {
                                            syn_empty_flag++;
                                        }
                                        else
                                        {
                                            syn_adherance_flag++;
                                        }
                                    }
                                }
                            }
                        }
                        if (syn_adherance_flag > 0) { syn_adherance_count++; } //cross patient count
                        if (syn_empty_flag > 0) { syn_empty_count++; }

                        if (!adherance_flag & !empty_flag & syn_adherance_flag == 0 & syn_empty_flag == 0) { missing_count = missing_count + 1; };

                        if (adherance_flag | ((!adherance_flag) & syn_adherance_flag == 1)) { auto_map_count++; }
                        if (!adherance_flag & syn_adherance_flag == 1) { auto_synonym_count++; }
                        if (!adherance_flag & syn_adherance_flag == 0) { no_map_count++; }
                    }

                    manual_map_count = range_count - auto_map_count - no_map_count;

                    double average_adherance = (double)adherance_count / range_count;
                    double average_empty = (double)empty_count / range_count;
                    double average_syn_adherance = (double)syn_adherance_count / range_count;
                    double average_syn_empty = (double)syn_empty_count / range_count;
                    double average_missing = (double)missing_count / range_count;
                    double average_auto_map = (double)auto_map_count / range_count;
                    double average_auto_exact = (double)auto_exact_count / range_count;
                    double average_auto_synonym = (double)auto_synonym_count / range_count;
                    double average_manual_map = (double)manual_map_count / range_count;
                    double average_no_map = (double)no_map_count / range_count;

                    ListViewGroup results_lvg = null;
                    if (dateRangeLabel != "")  // could be simplified as NOW groups are known before and (via dateRanges_listView)
                    {
                        bool group_exists = false;
                        foreach (ListViewGroup lvg in resultsLV.Groups)
                        {
                            if (lvg.Header == dateRangeLabel + " [n=" + range_count + "]")
                            {
                                group_exists = true;
                                results_lvg = lvg;
                                break;
                            }
                        }
                        if (!group_exists)
                        {
                            int results_lvg_index = resultsLV.Groups.Add(new ListViewGroup(dateRangeLabel + " [n=" + range_count + "]", HorizontalAlignment.Left));
                            results_lvg = resultsLV.Groups[results_lvg_index];
                        }
                    }



                    ListViewItem results_lvi = new ListViewItem(new string[] { dateRangeLabel, tmplt_structure_name, tmplt_structure_group, adherance_count.ToString(), average_adherance.ToString("P"), empty_count.ToString(), average_empty.ToString("P"), syn_adherance_count.ToString(), average_syn_adherance.ToString("P"), syn_empty_count.ToString(), average_syn_empty.ToString("P"), missing_count.ToString(), average_missing.ToString("P"), auto_map_count.ToString(), average_auto_map.ToString("P"), auto_exact_count.ToString(), average_auto_exact.ToString("P"), auto_synonym_count.ToString(), average_auto_synonym.ToString("P"), manual_map_count.ToString(), average_manual_map.ToString("P"), no_map_count.ToString(), average_no_map.ToString("P") });
                    if (dateRangeLabel != "") { results_lvi.Group = results_lvg; }
                    resultsLV.Items.Add(results_lvi);
                }
                //if (dateRangeLabel != "") { resultsLV.Groups.Add(default_lvg); } //add at end so 'default' group appears last

                resultsLV.EndUpdate();
            }
        }

        private void AnalyseNames(bool perDateRange, string dateRangeLabel, DateTime? rangeStart, DateTime? rangeEnd, ListView resultsLV)
        {

            if (matchedPlans.Count > 0 && Template_listView.Items.Count > 0)
            {
                //Results_textBox.Text = "";
                //Results_textBox.Text += "TEST" + Environment.NewLine;
                //foreach (this.MRN_listView.Items.)

                resultsLV.BeginUpdate();
                //ListViewGroup default_lvg = new ListViewGroup("Default (no group set)", HorizontalAlignment.Left);


                foreach (ListViewItem lvi in Template_listView.Items)  // iterate through structures in template
                {
                    string tmplt_structure_name = lvi.SubItems[0].Text;
                    string tmplt_structure_group = "";
                    if (lvi.SubItems.Count > 1)
                    {
                        tmplt_structure_group = lvi.SubItems[1].Text;
                    }
                    List<string> synonymes = new List<string>();
                    if (lvi.SubItems.Count > 2)
                    {
                        synonymes = lvi.SubItems[2].Text.Split(',').ToList();
                    }

                    int adherance_count = 0;
                    int syn_adherance_count = 0;
                    int empty_count = 0;
                    int syn_empty_count = 0;
                    int missing_count = 0;
                    int range_count = 0; // number of plans in specified date range

                    foreach (PlanPathAndStructures ppas in matchedPlans) // compare to valid plans
                    {
                        if (perDateRange)
                        {

                            if (!ppas.CourseStart.Value.IsInRange((DateTime)rangeStart, (DateTime)rangeEnd))
                            {
                                continue; // skip to next ppas
                            }
                            else { range_count = range_count + 1; }
                        }
                        else
                        {
                            range_count = matchedPlans.Count;
                        }

                        bool adherance_flag = false; // TRUE IF: template structure name match AND this contour is NOT empty
                        bool syn_adherance_flag = false; // TRUE IF: match to one of synonyms  AND this contour is NOT empty
                        bool empty_flag = false; // TRUE IF: template structure name match AND this contour is empty
                        bool syn_empty_flag = false; // TRUE IF: match to one of synonyms  AND this contour is empty

                        foreach (StructureInfo si in ppas.Structures)  // iterate over all structures of each valid plans
                        {

                            if (si.Id.Trim() == tmplt_structure_name.Trim()) //check for structure name match
                            {
                                if (si.isEmpty)
                                {
                                    empty_count = empty_count + 1;
                                    empty_flag = true;
                                }
                                else
                                {
                                    adherance_count = adherance_count + 1;
                                    adherance_flag = true;
                                }
                            }
                            else  //if no match check whether synonyms match
                            {
                                for (int i = 0; i < synonymes.Count; i++)  //iterate over synonyms of current template row
                                {
                                    if (si.Id == synonymes[i].Trim())
                                    {
                                        if (si.isEmpty)
                                        {
                                            syn_empty_flag = true;
                                            syn_empty_count = syn_empty_count + 1;
                                        }
                                        else
                                        {
                                            syn_adherance_flag = true;
                                            syn_adherance_count = syn_adherance_count + 1;
                                        }
                                        break; //if synonym match found no need to look further in synonyms list 
                                    }
                                }
                            }

                            if (adherance_flag | empty_flag | syn_adherance_flag | syn_empty_flag)
                            {
                                break; //if structure_name match of some sort (wheter synonym or empty) found no need to look further in same plan
                            }
                        }


                        if (!adherance_flag & !empty_flag & !syn_adherance_flag & !syn_empty_flag) { missing_count = missing_count + 1; };
                    }

                    double average_adherance = (double)adherance_count / range_count;
                    double average_empty = (double)empty_count / range_count;
                    double average_syn_adherance = (double)syn_adherance_count / range_count;
                    double average_syn_empty = (double)syn_empty_count / range_count;
                    double average_missing = (double)missing_count / range_count;

                    ListViewGroup results_lvg = null;
                    if (dateRangeLabel != "")  // could be simplified as NOW groups are known before and (via dateRanges_listView)
                    {
                        bool group_exists = false;
                        foreach (ListViewGroup lvg in resultsLV.Groups)
                        {
                            if (lvg.Header == dateRangeLabel + " [n=" + range_count + "]")
                            {
                                group_exists = true;
                                results_lvg = lvg;
                                break;
                            }
                        }
                        if (!group_exists)
                        {
                            int results_lvg_index = resultsLV.Groups.Add(new ListViewGroup(dateRangeLabel + " [n=" + range_count + "]", HorizontalAlignment.Left));
                            results_lvg = resultsLV.Groups[results_lvg_index];
                        }
                    }



                    ListViewItem results_lvi = new ListViewItem(new string[] { dateRangeLabel, tmplt_structure_name, tmplt_structure_group, adherance_count.ToString(), average_adherance.ToString("P"), empty_count.ToString(), average_empty.ToString("P"), syn_adherance_count.ToString(), average_syn_adherance.ToString("P"), syn_empty_count.ToString(), average_syn_empty.ToString("P"), missing_count.ToString(), average_missing.ToString("P") });
                    if (dateRangeLabel != "") { results_lvi.Group = results_lvg; }
                    resultsLV.Items.Add(results_lvi);
                }
                //if (dateRangeLabel != "") { resultsLV.Groups.Add(default_lvg); } //add at end so 'default' group appears last

                resultsLV.EndUpdate();
            }
        }

        private void LoadTemplate_button_Click(object sender, EventArgs e)
        {
            try
            {
                System.Windows.Forms.OpenFileDialog OpenFileDialog = new System.Windows.Forms.OpenFileDialog();
                if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\templates"))
                {
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\data\templates");
                }
                OpenFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\templates";
                OpenFileDialog.Filter = "Tab delimited text files (*.txt)|*.txt|All files (*.*)|*.*";
                OpenFileDialog.RestoreDirectory = true;

                if (OpenFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Template_listView.Items.Clear();
                    var lines = File.ReadAllLines(OpenFileDialog.FileName);
                    foreach (string line in lines)
                    {
                        var parts = line.Split('\t');
                        ListViewItem lvi = new ListViewItem(parts[0]);
                        if (parts.Count() > 1)
                        {
                            lvi.SubItems.Add(parts[1]); //add 'Group' item
                            if (parts.Count() > 2)
                            {
                                lvi.SubItems.Add(parts[2]); //add 'Synonyms' item
                            }

                        }
                        Template_listView.Items.Add(lvi);
                    }
                }
            }
            catch (Exception errorMsg)
            {
                MessageBox.Show(errorMsg.Message, "Error reading a file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void Export_button_Click(object sender, EventArgs e)
        {
            try
            {
                //string fp = OutputDir + "\\test.csv";
                //ListViewToCSV.DoListViewToCSV(Results_listView, fp, true);
                //MessageBox.Show("Results exported as tab delimited text file to: " + fp + ".");

                SaveFileDialog dlg = new SaveFileDialog();
                dlg.DefaultExt = "csv";
                dlg.Filter = "Tab delimited text (*.csv)|*.csv|All files (*.*)|*.*";

                if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\results"))
                {
                    dlg.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\results";
                }
                else { dlg.InitialDirectory = @"C:\temp\StructureNameAnalyser\data\results"; }

                DialogResult result = dlg.ShowDialog();

                string fn;
                if (result == DialogResult.OK)
                {
                    // Save document
                    fn = dlg.FileName;
                    ListViewToCSV.DoListViewToCSV(Results_listView, fn, true);
                    MessageBox.Show("Results exported as tab delimited text file to:"+ Environment.NewLine + fn, "Export successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception errorMsg)
            {
                MessageBox.Show("Error during export! See detailed error message in brackets." + Environment.NewLine + "(" + errorMsg.Message + ")", "Export error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void validDose_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            validDose_panel.Enabled = (sender as CheckBox).Checked;
        }

        private void advFilters_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            advFilters_panel.Visible = advFilters_checkBox.Checked;
            advFilters_panel2.Visible = advFilters_checkBox.Checked;
            advBrCaFilters_panel.Visible = advFilters_checkBox.Checked;
        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void Results_listView_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void analyseFrequency_button_Click(object sender, EventArgs e)
        {
            if (matchedPlans.Count > 0)
            {
                StringComparer sc;
                if (analyseFrequCaseSens_radioButton.Checked)
                {
                    sc = StringComparer.CurrentCulture;
                } else
                {
                    sc = StringComparer.CurrentCultureIgnoreCase;
                }
                //Dictionary<string, int> structureIdHistogram = new Dictionary<string, int>(sc);
                Dictionary<string, FaEntry> structureIdHistogram = new Dictionary<string, FaEntry>(sc);
                frequency_listView.Items.Clear();
                analyseFrequUnique_label.Text = "";

                // Volume variables
                Boolean FAvol_flag = true;
                Double FAlowerVol = 0;
                Double FAupperVol = 0;
                Double FAthreshVol = 0;
                try
                { 
                FAlowerVol = Convert.ToDouble(FAlowerVol_textBox.Text.Trim());
                FAupperVol = Convert.ToDouble(FAupperVol_textBox.Text.Trim());
                FAthreshVol = Convert.ToDouble(FAthreshVol_textBox.Text.Trim())/100;
                }
                catch { FAvol_flag = false; }

                // Dose variables
                int FAdose_flag = 0;
                if (dMean_radioButton.Checked) { FAdose_flag = 1; }
                if (D95_radioButton.Checked) { FAdose_flag = 2; }
                Double FAlowerDose = 0;
                Double FAupperDose = 0;
                Double FAthreshDose = 0;
                try
                {
                    FAlowerDose = Convert.ToDouble(FAlowerDmean_textBox.Text.Trim());
                    FAupperDose = Convert.ToDouble(FAupperDmean_textBox.Text.Trim());
                    FAthreshDose = Convert.ToDouble(FAthreshDmean_textBox.Text.Trim()) / 100;
                }
                catch { FAdose_flag = 0; }

                if (!FAvol_flag && !(FAlowerVol_textBox.Text == "" && FAupperVol_textBox.Text == "" && FAthreshVol_textBox.Text == "")) { MessageBox.Show("Error in 'Volume' match constraints. They won't be applied.", "Error in Vol. match constraints", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                if (FAdose_flag == 0 && !(FAlowerDmean_textBox.Text == "" && FAupperDmean_textBox.Text == "" && FAthreshDmean_textBox.Text == "")) { MessageBox.Show("Error in 'Dose' match constraints. They won't be applied.", "Error in Vol. match constraints", MessageBoxButtons.OK, MessageBoxIcon.Error); }

                        foreach (PlanPathAndStructures ppas in matchedPlans)
                {
                    foreach (StructureInfo si in ppas.Structures)
                    {
                        // add the structure id to the histogram or increase the usage count if already there.
                        //int count = 0;
                        //if (structureIdHistogram.TryGetValue(si.Id, out count))
                        if (structureIdHistogram.ContainsKey(si.Id))
                        {
                            //structureIdHistogram[si.Id] = count + 1;
                            structureIdHistogram[si.Id].occurence++;
                        }
                        else
                        {
                            structureIdHistogram.Add(si.Id, new FaEntry(1, 0, 0, 100000000, 0, 0, 0, 100000000, 0, 0, 0, 100000000, 0, 0, 0));
                            //if (FAvol_flag && !si.isEmpty && si.vol >= FAlowerVol && si.vol <= FAupperVol)
                            //{
                            //    structureIdHistogram[si.Id].volMatch++;
                            //}
                            //if (FAdmean_flag && !si.isEmpty && si.dMean >= FAlowerDmean && si.dMean <= FAupperDmean) { structureIdHistogram[si.Id].dmeanMatch++; }
                            //if (!si.isEmpty && si.vol >= FAlowerVol && si.vol <= FAupperVol) { structureIdHistogram[si.Id].volMatch++; }
                        }
                        // Add dose/vol details to current si
                        if (si.isEmpty) { structureIdHistogram[si.Id].empty++; }
                        else
                        {
                            structureIdHistogram[si.Id].volCum = structureIdHistogram[si.Id].volCum + si.vol;
                            if (structureIdHistogram[si.Id].volMin > si.vol) { structureIdHistogram[si.Id].volMin = si.vol; }
                            if (structureIdHistogram[si.Id].volMax < si.vol) { structureIdHistogram[si.Id].volMax = si.vol; }
                            if (Double.IsNaN(si.dMean)) { structureIdHistogram[si.Id].dmeanNan++; }
                            if (!Double.IsNaN(si.dMean)) { structureIdHistogram[si.Id].dmeanCum = structureIdHistogram[si.Id].dmeanCum + si.dMean; }
                            if (structureIdHistogram[si.Id].dmeanMin > si.dMean) { structureIdHistogram[si.Id].dmeanMin = si.dMean; }
                            if (structureIdHistogram[si.Id].dmeanMax < si.dMean) { structureIdHistogram[si.Id].dmeanMax = si.dMean; }
                            if (Double.IsNaN(si.d95)) { structureIdHistogram[si.Id].d95Nan++; }
                            if (!Double.IsNaN(si.d95)) { structureIdHistogram[si.Id].d95Cum = structureIdHistogram[si.Id].d95Cum + si.d95; }
                            if (structureIdHistogram[si.Id].d95Min > si.d95) { structureIdHistogram[si.Id].d95Min = si.d95; }
                            if (structureIdHistogram[si.Id].d95Max < si.d95) { structureIdHistogram[si.Id].d95Max = si.d95; }
                        }
                        


                        if (FAvol_flag && !si.isEmpty && !(Double.IsNaN(si.vol)) && si.vol >= FAlowerVol && si.vol <= FAupperVol) 
                        {
                            structureIdHistogram[si.Id].volMatch++;
                        }
                        if (FAdose_flag != 0 && !si.isEmpty && !(Double.IsNaN(si.dMean)) && si.dMean >= FAlowerDose && si.dMean <= FAupperDose)
                        {
                            structureIdHistogram[si.Id].doseMatch++;
                        }
                    }



                }


                //List<KeyValuePair<string, int>> structureIdHistogramList = structureIdHistogram.ToList();

                List<KeyValuePair<string, int>> structureIdHistogramList = new List<KeyValuePair<string, int>>();
                foreach (KeyValuePair<string, FaEntry> entry in structureIdHistogram)
                {
                    Boolean FAvolMatch_flag = false;
                    if (FAvol_flag && ((Double)entry.Value.volMatch / entry.Value.occurence) >= FAthreshVol) { FAvolMatch_flag = true; }
                    Boolean FAdmeanMatch_flag = false;
                    if (FAdose_flag != 0 && ((Double)entry.Value.doseMatch / entry.Value.occurence) >= FAthreshDose) { FAdmeanMatch_flag = true; }

                    if (!FAvol_flag && FAdose_flag == 0 ) //no vol/dose filtering
                    {
                        structureIdHistogramList.Add(new KeyValuePair<string, int>(entry.Key, entry.Value.occurence));
                    }
                    else //at least one dose/vol filter properly set
                    {
                        
                        //if ((FAvol_flag && (Double) entry.Value.volMatch/entry.Value.occurence >= FAthreshVol) | (FAdmean_flag && (Double)entry.Value.dmeanMatch / entry.Value.occurence >= FAthreshDmean))
                        if (FAvol_flag && FAdose_flag != 0)
                        {
                            if (FAvolMatch_flag && FAdmeanMatch_flag)
                            {
                                structureIdHistogramList.Add(new KeyValuePair<string, int>(entry.Key, entry.Value.occurence));
                            }
                        }
                        else if (FAvol_flag && FAdose_flag == 0)
                        {
                            if (FAvolMatch_flag)
                            {
                                structureIdHistogramList.Add(new KeyValuePair<string, int>(entry.Key, entry.Value.occurence));
                            }
                        }
                        else if (!FAvol_flag && FAdose_flag != 0)
                        {
                            if (FAdmeanMatch_flag)
                            {
                                structureIdHistogramList.Add(new KeyValuePair<string, int>(entry.Key, entry.Value.occurence));
                            }
                        }


                    }
                }

                
                // sort to put the structures with highest frequency usage first.
                structureIdHistogramList.Sort((firstPair, nextPair) =>
                {
                    return nextPair.Value.CompareTo(firstPair.Value);
                }
                );

                //Console.WriteLine(Environment.NewLine + "** Structure Name Frequency Analysis **" + Environment.NewLine);
                try
                {
                    Regex filter_regex = new Regex(@frequencyRegEx_textBox.Text, RegexOptions.IgnoreCase);
                    frequency_listView.BeginUpdate();
                    foreach (var kvp in structureIdHistogramList)
                    {
                        //Console.WriteLine(string.Format("{0},{1}", kvp.Key, kvp.Value));
                        bool regex_flag = false;
                        if (frequencyRegEx_textBox.Text != "")
                        {
                            regex_flag = filter_regex.IsMatch(kvp.Key);
                        }
                        else
                        {
                            regex_flag = true; //if no regex set in GUI don't use regex condition
                        }
                        if (regex_flag && kvp.Value >= frequency_numericUpDown.Value)
                        {
                            ListViewItem lvi = new ListViewItem(kvp.Key);
                            lvi.Name = kvp.Key; // to make it findable via search button
                            lvi.SubItems.Add(kvp.Value.ToString());
                            if (structureIdHistogram.ContainsKey(kvp.Key))
                            {
                                lvi.SubItems.Add(structureIdHistogram[kvp.Key].empty.ToString());
                                if (structureIdHistogram[kvp.Key].volMatch > 0) { lvi.SubItems.Add(structureIdHistogram[kvp.Key].volMatch.ToString()); } else { lvi.SubItems.Add(""); }
                                if (structureIdHistogram[kvp.Key].doseMatch > 0) { lvi.SubItems.Add(structureIdHistogram[kvp.Key].doseMatch.ToString()); } else { lvi.SubItems.Add(""); }
                                lvi.SubItems.Add(String.Format("{0:0.00}", structureIdHistogram[kvp.Key].volCum / (structureIdHistogram[kvp.Key].occurence - structureIdHistogram[kvp.Key].empty)));
                                lvi.SubItems.Add(String.Format("{0:0.0}", structureIdHistogram[kvp.Key].volMin));
                                lvi.SubItems.Add(String.Format("{0:0.0}", structureIdHistogram[kvp.Key].volMax));
                                lvi.SubItems.Add(String.Format("{0:0.00}", structureIdHistogram[kvp.Key].dmeanCum / (structureIdHistogram[kvp.Key].occurence - structureIdHistogram[kvp.Key].empty - structureIdHistogram[kvp.Key].dmeanNan)));
                                lvi.SubItems.Add(String.Format("{0:0.0}", structureIdHistogram[kvp.Key].dmeanMin));
                                lvi.SubItems.Add(String.Format("{0:0.0}", structureIdHistogram[kvp.Key].dmeanMax));
                                lvi.SubItems.Add(String.Format("{0:0.00}", structureIdHistogram[kvp.Key].d95Cum / (structureIdHistogram[kvp.Key].occurence - structureIdHistogram[kvp.Key].empty - structureIdHistogram[kvp.Key].d95Nan)));
                                lvi.SubItems.Add(String.Format("{0:0.0}", structureIdHistogram[kvp.Key].d95Min));
                                lvi.SubItems.Add(String.Format("{0:0.0}", structureIdHistogram[kvp.Key].d95Max));
                            }
                            if (buildingTemplate.allMembersStem.Contains(lvi.Text))
                            {
                                lvi.ForeColor = Color.Red;
                                lvi.Checked = false;
                                if (!hideUsed_checkBox.Checked) { frequency_listView.Items.Add(lvi); }
                            }
                            else
                            {
                                frequency_listView.Items.Add(lvi);
                            }


                        }
                        //frequency_listView.Items.Add(new ListViewItem(new string[] { kvp.Key, kvp.Value.ToString() }));
                    }
                    // autosize the listView columns to content
                    frequency_listView.Columns[0].Width = -2;
                    frequency_listView.Columns[1].Width = -2;
                    frequency_listView.EndUpdate();
                    analyseFrequUnique_label.Text = "n = " + frequency_listView.Items.Count.ToString();
                }
                catch
                {
                    MessageBox.Show("Likely problem with Filter Regular Expression. Please revise", "Regex problem", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void label2_Click_1(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void frequencyRegtextBox_TextChanged(object sender, EventArgs e)
        {

        }


        private void setGroup_button_Click(object sender, EventArgs e)
        {
            if (frequency_listView.SelectedItems.Count > 0)
            {
                groupName_textBox.Text = frequency_listView.SelectedItems[0].Text;
            }

        }

        private void clearGroup_button_Click(object sender, EventArgs e)
        {
            groupName_textBox.Text = "";
        }


        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        }

        private void addRow_button_Click(object sender, EventArgs e)
        {
            if (frequency_listView.CheckedItems.Count > 0)
            {

                if (LSSNdictV1_radioButton.Checked && forHisto_radioButton.Checked)
                {
                    foreach (ListViewItem lvi in frequency_listView.CheckedItems)
                    {
                        template_richTextBox.Text += lvi.Text + "\t" + groupName_textBox.Text + Environment.NewLine;
                        lvi.ForeColor = SystemColors.InactiveCaption;
                        lvi.Checked = false;
                    }
                }
                else
                {
                    if (frequency_listView.SelectedItems.Count == 0)
                    {
                        //TODO
                        MessageBox.Show("No row selected in frequencies list", "Select one determine the name of synonym group.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        string synonyms = "";
                        string synonym_group_name = "";
                        foreach (ListViewItem lvi in frequency_listView.CheckedItems)
                        {
                            if (lvi.Selected)
                            {
                                synonym_group_name = lvi.Text;
                            }
                            else
                            {
                                synonyms += lvi.Text + ", ";
                            }
                            lvi.ForeColor = SystemColors.InactiveCaption;
                            lvi.Checked = false;
                        }
                        synonyms = synonyms.TrimEnd(new char[] { ',', ' ' }); //remove last ', '
                        template_richTextBox.Text += synonym_group_name + "\t" + groupName_textBox.Text + "\t" + synonyms + Environment.NewLine;
                    }
                }





            }
        }

        private void exportTemplate_button_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.DefaultExt = "txt";
            dlg.Filter = "Tab delimited text (*.txt)|*.txt|All files (*.*)|*.*";

            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\templates"))
            {
                dlg.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\templates";
            }
            else { dlg.InitialDirectory = @"C:\temp\StructureNameAnalyser\data\templates"; }

            DialogResult result = dlg.ShowDialog();

            if (result == DialogResult.OK)
            {
                // Save document
                string fn = dlg.FileName;
                File.WriteAllText(fn, template_richTextBox.Text);
            }
        }

        private void label20_Click(object sender, EventArgs e)
        {

        }



        private void forHisto_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            //groupName_textBox.Text = "";
            //if (forHisto_radioButton.Checked)
            //{
            //    groupName_textBox.ReadOnly = true;
            //}
            //else
            //{
            //    groupName_textBox.ReadOnly = false;
            //}
        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void label25_Click(object sender, EventArgs e)
        {

        }

        private void addDateRangeEntry_button_Click(object sender, EventArgs e)
        {
            dateRanges_listView.Items.Add(new ListViewItem(new string[] { dateRangeLabel_textBox.Text, startDate_dateTimePicker.Text, endDate_dateTimePicker.Text }));
        }

        private void clearDateRangeEntries_button_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem lvi in dateRanges_listView.CheckedItems)
            {
                dateRanges_listView.Items.Remove(lvi);
            }
        }

        private void dateRanges_listView_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void Export2_button_Click(object sender, EventArgs e)
        {

        }

        private void Analyse2_button_Click(object sender, EventArgs e)
        {
            Matrix_listView.Clear();
            Matrix_listView.Sorting = SortOrder.None;
            Matrix_listView.BeginUpdate();

            //Add columns with variable width
            Matrix_listView.Columns.Add("MRN", 60);
            Matrix_listView.Columns.Add("Course", 100);
            Matrix_listView.Columns.Add("Plan", 100);
            Matrix_listView.Columns.Add("Course Start", 70);
            foreach (ListViewItem lvi in RepTemp_objectListView.Items)
            {
                Matrix_listView.Columns.Add(lvi.Text, 60);  //'-2' is autosize
            }
            Matrix_listView.Columns.Add(" ",1);



            //Fill MRNs into rows of first column
            ListViewItem temp_new_lvi;
            foreach (PlanPathAndStructures ppas in matchedPlans)
            {
                temp_new_lvi = Matrix_listView.Items.Add(new ListViewItem(new string[] { ppas.MRN }));
                temp_new_lvi.Name = ppas.MRN;
                temp_new_lvi.SubItems.Add(ppas.Course);
                temp_new_lvi.SubItems.Add(ppas.Plan);
                if (ppas.CourseStart != null)
                {
                    DateTime temp_date = (DateTime)ppas.CourseStart;
                    temp_new_lvi.SubItems.Add(temp_date.ToString("d"));
                }
                else
                {
                    temp_new_lvi.SubItems.Add("");
                }
            }
            temp_new_lvi = Matrix_listView.Items.Add("*MAPPING*");
            temp_new_lvi.UseItemStyleForSubItems = false;
            temp_new_lvi.SubItems.Add("");
            temp_new_lvi.SubItems.Add("");
            temp_new_lvi.SubItems.Add(new DateTime(2200, 1, 1).ToString("d"), Color.White, Color.White, this.Font);
            temp_new_lvi = Matrix_listView.Items.Add(">> Auto");
            temp_new_lvi.UseItemStyleForSubItems = false;
            temp_new_lvi.SubItems.Add("");
            temp_new_lvi.SubItems.Add("");
            temp_new_lvi.SubItems.Add(new DateTime(2200, 1, 2).ToString("d"), Color.White, Color.White, this.Font);
            temp_new_lvi = Matrix_listView.Items.Add(">> Manual");
            temp_new_lvi.UseItemStyleForSubItems = false;
            temp_new_lvi.SubItems.Add("");
            temp_new_lvi.SubItems.Add("");
            temp_new_lvi.SubItems.Add(new DateTime(2200, 1, 3).ToString("d"), Color.White, Color.White, this.Font);
            temp_new_lvi = Matrix_listView.Items.Add(">> None");
            temp_new_lvi.UseItemStyleForSubItems = false;
            temp_new_lvi.SubItems.Add("");
            temp_new_lvi.SubItems.Add("");
            temp_new_lvi.SubItems.Add(new DateTime(2200, 1, 4).ToString("d"), Color.White, Color.White, this.Font);

            //Fill rows of remaining columns one at a time
            foreach (ListViewItem lvi in RepTemp_objectListView.Items)
            {
                string structure_name = lvi.Text;
                //List<string> synonyms = new List<string>();
                Dictionary<string, string> synonyms_with_details = new Dictionary<string, string>();

                if (lvi.SubItems.Count > 2)
                {
                    //synonyms = lvi.SubItems[2].Text.Split(',').ToList();
                    List<string> temp_synonyms_with_details = new List<string>();
                    temp_synonyms_with_details = lvi.SubItems[2].Text.Split(',').ToList();
                    foreach (string syn_details in temp_synonyms_with_details)
                    {
                        if (syn_details.Contains("{"))
                        {
                            string temp_syn = "";
                            string temp_details = "";
                            temp_syn = syn_details.Split('{')[0].Trim();
                            temp_details = syn_details.Split('{')[1].Trim(new char[] { '}', ' ' });
                            //Console.WriteLine("syn: '" + temp_syn + "'; " + "details: '" + temp_details + "'");
                            synonyms_with_details.Add(temp_syn, temp_details);
                        }
                        else
                        {
                            synonyms_with_details.Add(syn_details, null);
                        }
                    }
                }

                int auto_mapping_count = 0;
                int manual_mapping_count = 0;
                int non_mapping_count = 0;

                //for (int i = 0; i < synonyms.Count(); i++)
                //{
                //Matrix_listView.Items.Add(new ListViewItem(new string[] { synonyms[i].Trim() }));
                //}

                int ppas_iterator = 0;

                foreach (PlanPathAndStructures ppas in matchedPlans) // compare to valid plans
                {

                    int struc_count = 0;
                    int struc_empty_count = 0;
                    int synonyms_count = 0;
                    int synonyms_empty_count = 0;

                    foreach (StructureInfo si in ppas.Structures)
                    {
                        //structure name
                        if (structure_name == si.Id)
                        {
                            if (si.isEmpty)
                            {
                                struc_empty_count = struc_empty_count + 1;
                            }
                            else
                            {
                                struc_count = struc_count + 1;
                            }
                        }
                        //synonyms
                        for (int i = 0; i < synonyms_with_details.Count(); i++)
                        {
                            if (synonyms_with_details.Keys.ElementAt(i).Trim() == si.Id)
                            {
                                if (synonyms_with_details.Values.ElementAt(i) != null)
                                {
                                    bool isConditionMet = false;
                                    var jintEval = new Engine()  //check details expression via JINT (needs to evaluate to true or false)
                                    .SetValue("Vol", si.vol)
                                    .SetValue("Dmean", si.dMean)
                                    .SetValue("D95", si.d95)
                                    .Execute(synonyms_with_details.Values.ElementAt(i)) // details expression of current synonym
                                    .GetCompletionValue()
                                    .ToObject();
                                    isConditionMet = (bool)jintEval;

                                    if (isConditionMet)
                                    {
                                        if (si.isEmpty)
                                        {
                                            synonyms_empty_count = synonyms_empty_count + 1;
                                        }
                                        else
                                        {
                                            synonyms_count = synonyms_count + 1;
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("Matrix - MRN '" + ppas.MRN + "', Structure '" + si.Id + "' (Vol: "
                                         + si.vol.ToString() + "; Dmean: "
                                         + si.dMean.ToString() + ")  does not meet: '"
                                         + synonyms_with_details.Values.ElementAt(i) + "'");
                                    }
                                }
                                else
                                {
                                    if (si.isEmpty)
                                    {
                                        synonyms_empty_count = synonyms_empty_count + 1;
                                    }
                                    else
                                    {
                                        synonyms_count = synonyms_count + 1;
                                    }
                                }
                            }
                        }
                    }
                    //display
                    string struc_empty_string = "";
                    if (struc_empty_count > 0) { struc_empty_string = " (" + struc_empty_count.ToString() + ")"; }
                    string synonyms_empty_string = "";
                    if (synonyms_empty_count > 0) { synonyms_empty_string = " (" + synonyms_empty_count.ToString() + ")"; }
                    string cell_content = struc_count.ToString() + struc_empty_string + " / " + synonyms_count.ToString() + synonyms_empty_string;
                    Color cell_background_color = Color.Empty;
                    //if (cell_content == "1 / 0") { cell_background_color = Color.Green; }
                    //if (cell_content == "0 / 1") { cell_background_color = Color.Orange; }
                    //if (cell_background_color == Color.Empty) { cell_background_color = Color.Red; }
                    if (struc_count == 1)
                    {
                        cell_background_color = Color.Green;
                        auto_mapping_count++;
                    }
                    else
                    {
                        if (synonyms_count == 1)
                        {
                            cell_background_color = Color.Orange;
                            auto_mapping_count++;
                        }
                        if (synonyms_count > 1)
                        {
                            cell_background_color = Color.Yellow;
                            manual_mapping_count++;
                        }
                    }
                    if (struc_count == 0 && synonyms_count == 0)
                    {
                        cell_background_color = Color.Red;
                        non_mapping_count++;
                    }
                    if (struc_empty_count > 0 || synonyms_empty_count > 0)
                    {
                        cell_background_color = Color.LightBlue;
                    }

                    //if (cell_background_color == Color.Empty) { cell_background_color = Color.Red; }
                    Matrix_listView.Items[ppas_iterator].UseItemStyleForSubItems = false;
                    Matrix_listView.Items[ppas_iterator].SubItems.Add(cell_content, Color.Black, cell_background_color, this.Font);

                    ppas_iterator = ppas_iterator + 1;
                }

                double percent_auto = (double)auto_mapping_count / ppas_iterator*100;
                double percent_manual = (double)manual_mapping_count / ppas_iterator*100;
                double percent_none = (double)non_mapping_count / ppas_iterator*100;
                Matrix_listView.Items[ppas_iterator].SubItems.Add("");
                Matrix_listView.Items[ppas_iterator + 1].SubItems.Add(percent_auto.ToString("0") + "% (" + auto_mapping_count.ToString() + ")");
                Matrix_listView.Items[ppas_iterator + 2].SubItems.Add(percent_manual.ToString("0") + "% (" + manual_mapping_count.ToString() + ")");
                Matrix_listView.Items[ppas_iterator + 3].SubItems.Add(percent_none.ToString("0") + "% (" + non_mapping_count.ToString() + ")");
                // display
                //foreach (ListViewItem matrix_row in Matrix_listView.Items)
                //{
                //    matrix_row.SubItems.Add(pref_count.ToString() + " / " + synonyms_count.ToString());
                //}
            }


            //foreach (ListViewItem lvi in Template_listView.Items)  // iterate through structures in template
            //{
            //    string tmplt_structure_name = lvi.SubItems[0].Text;
            //    string tmplt_structure_group = "";
            //    if (lvi.SubItems[1] != null)
            //    {
            //        tmplt_structure_group = lvi.SubItems[1].Text;
            //    }
            //    int adherance_count = 0;
            //    int empty_count = 0;
            //    int missing_count = 0;
            //    int range_count = 0; // number of plans in specified date range

            //    foreach (PlanPathAndStructures ppas in matchedPlans) // compare to valid plans
            //    {
            //        if (perDateRange)
            //        {

            //            if (!ppas.CourseStart.Value.IsInRange((DateTime)rangeStart, (DateTime)rangeEnd))
            //            {
            //                continue; // skip to next ppas
            //            }
            //            else { range_count = range_count + 1; }
            //        }
            //        else
            //        {
            //            range_count = matchedPlans.Count;
            //        }

            //        bool adherance_flag = false; // TRUE IF: names match AND contour is not empty
            //        bool empty_flag = false; // TRUE IF: names match AND contour is empty
            //        foreach (StructureInfo si in ppas.Structures)
            //        {
            //            if (tmplt_structure_name == si.Id)
            //            {
            //                if (si.isEmpty)
            //                {
            //                    empty_count = empty_count + 1;
            //                    empty_flag = true;
            //                }
            //                else
            //                {
            //                    adherance_count = adherance_count + 1;
            //                    adherance_flag = true;
            //                    break; //if adherance found no need to look further 
            //                }

            //            }
            //        }
            //        if (!adherance_flag & !empty_flag) { missing_count = missing_count + 1; };
            //    }


            //Matrix_listView.S
            //Matrix_listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            //Matrix_listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
            Matrix_listView.EndUpdate();
        }

        private void groupByDate_checkBox_CheckedChanged_1(object sender, EventArgs e)
        {
            groupByDate_panel.Enabled = groupByDate_checkBox.Checked;
        }

        private void label35_Click(object sender, EventArgs e)
        {

        }

        private void splitContainer2_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Matches_listView.Visible)
            {
                Matches_listView.Visible = false;
                currentPlanInTree_linkLabel.Visible = false;
            }
            else
            {
                Matches_listView.Visible = true;
                currentPlanInTree_linkLabel.Visible = true;
            }
        }

        private void Matrix_listView_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            //System.Text.StringBuilder messageBoxCS = new System.Text.StringBuilder();
            //messageBoxCS.AppendFormat("{0} = {1}", "IsSelected", e.IsSelected);
            //messageBoxCS.AppendLine();
            //messageBoxCS.AppendFormat("{0} = {1}", "Item", e.Item);
            //messageBoxCS.AppendLine();
            //messageBoxCS.AppendFormat("{0} = {1}", "ItemIndex", e.ItemIndex);
            //messageBoxCS.AppendLine();
            //MessageBox.Show(messageBoxCS.ToString(), "ItemSelectionChanged Event");

            if (e.IsSelected)
            {
                if (Regex.IsMatch(e.Item.Text, @"^\d+$")) //avoid if clicking on summary fields without mrn in first column
                {

                    MappingDetailsOLVmodel.Clear();

                    mappingDetails_label2.Text =    "MRN: " + e.Item.Text + Environment.NewLine +
                                                    "Course: " + e.Item.SubItems[1].Text + Environment.NewLine +
                                                    "Plan: " + e.Item.SubItems[2].Text + Environment.NewLine +
                                                    "Course Start: " + e.Item.SubItems[3].Text;

                    List <StructureInfo> temp_structures_list = Matches_listView.Items.Find(e.Item.Text, false).FirstOrDefault().Tag as List<StructureInfo>;
                    if (temp_structures_list != null)
                    {

                        string ms;
                        string tn;
                        int tnc;
                        string gn = "REMAINDER (matches exact or via single synonym and non-matches)";

                        //Go through all structures of current plan and attempt matching to template
                        foreach (StructureInfo st_iterator in temp_structures_list)
                        {
                            ms = "";
                            tn = "";
                            tnc = 0;


                            foreach (ListViewItem lvi in RepTemp_objectListView.Items)
                            {
                                Boolean leave_nested_loop = false;
                                string template_name = lvi.Text;
                                //List<string> synonyms = new List<string>();
                                Dictionary<string, string> synonyms_with_details = new Dictionary<string, string>();
                                        
                                if (lvi.SubItems.Count > 2)
                                {
                                    //synonyms = lvi.SubItems[2].Text.Split(',').ToList();
                                    List<string> temp_synonyms_with_details = new List<string>();
                                    temp_synonyms_with_details = lvi.SubItems[2].Text.Split(',').ToList();
                                    foreach (string syn_details in temp_synonyms_with_details)
                                    {
                                        if (syn_details.Contains("{"))
                                        {
                                            string temp_syn = "";
                                            string temp_details = "";
                                            temp_syn = syn_details.Split('{')[0].Trim();
                                            temp_details = syn_details.Split('{')[1].Trim(new char[] { '}', ' ' });
                                            //Console.WriteLine("syn: '" + temp_syn + "'; " + "details: '" + temp_details + "'");
                                            synonyms_with_details.Add(temp_syn, temp_details);
                                        }
                                        else
                                        {
                                            synonyms_with_details.Add(syn_details, null);
                                        }
                                    }
                                }

                                //structure name
                                if (template_name == st_iterator.Id)
                                {

                                    if (st_iterator.isEmpty)
                                    {
                                        ms = "EMPTY (name match but empty)";
                                    }
                                    else
                                    {
                                        ms = "NAME (exact name match)";
                                    }
                                    tn = template_name;
                                    break; //if name match don't look for synonyms
                                }
                                //synonyms
                                for (int i = 0; i < synonyms_with_details.Count(); i++)
                                {
                                    if (synonyms_with_details.Keys.ElementAt(i).Trim() == st_iterator.Id)
                                    {
                                        if (synonyms_with_details.Values.ElementAt(i) != null)
                                        {
                                            bool isConditionMet = false;
                                            var jintEval = new Engine()  //check details expression via JINT (needs to evaluate to true or false)
                                            .SetValue("Vol", st_iterator.vol)
                                            .SetValue("Dmean", st_iterator.dMean)
                                            .SetValue("D95", st_iterator.d95)
                                            .Execute(synonyms_with_details.Values.ElementAt(i)) // details expression of current synonym
                                            .GetCompletionValue()
                                            .ToObject();
                                            isConditionMet = (bool)jintEval;

                                            if (isConditionMet)
                                            {
                                                if (st_iterator.isEmpty)
                                                {
                                                    ms = "EMPTY (syonym match but empty)";
                                                }
                                                else
                                                {
                                                    ms = "SYNONYM (synonym match)";
                                                }
                                                tn = template_name;
                                                leave_nested_loop = true; //if synonym match don't look for further synonyms => leave nested loop 
                                                break;
                                            }
                                            else
                                            {
                                                Console.WriteLine("Single Details - MRN '" + e.Item.Text + "', Structure '" + st_iterator.Id + "' (Vol: "
                                                 + st_iterator.vol.ToString() + "; Dmean: "
                                                 + st_iterator.dMean.ToString() + ")  does not meet: '"
                                                 + synonyms_with_details.Values.ElementAt(i) + "'");
                                            }
                                        }
                                        else
                                        {
                                            if (st_iterator.isEmpty)
                                            {
                                                ms = "EMPTY (syonym match but empty)";
                                            }
                                            else
                                            {
                                                ms = "SYNONYM (synonym match)";
                                            }
                                            tn = template_name;
                                            leave_nested_loop = true; //if synonym match don't look for further synonyms => leave nested loop 
                                            break;
                                        }
                                        

                                    }
                                }

                                if (leave_nested_loop)
                                {
                                    break;
                                }
                            }

                            if (ms == "") //i.e. no exact or synonym match
                            {
                                ms = "NO (matching template name)";
                                tnc = 0;
                            }
                            else
                            {
                                tnc = 1; //if >1 will be fixed further down 
                            }

                            MappingDetailsOLVmodel.Add(new OLV_MappingDetails(MappingDetailsOLVmodel.Count + 1, e.Item.Text, e.Item.SubItems[1].Text, e.Item.SubItems[2].Text, st_iterator.Id, st_iterator.vol, st_iterator.d95, st_iterator.dMean, st_iterator.dMedian, ms, tn, tnc, gn));
                        }

                        //Determine template_names that couldn't be matched or where there are muliple matches (i.e. via multiple synonyms in absence of exact match)
                        foreach (ListViewItem lvi in RepTemp_objectListView.Items)
                        {
                            int matched = 0;
                            foreach (OLV_MappingDetails md in MappingDetailsOLVmodel)
                            {
                                if (lvi.Text == md.templateName)
                                {
                                    matched++;
                                }
                            }

                            if (matched == 0)  // no matches
                            {
                                tn = lvi.Text;
                                MappingDetailsOLVmodel.Add(new OLV_MappingDetails(MappingDetailsOLVmodel.Count + 1, e.Item.Text, e.Item.SubItems[1].Text, e.Item.SubItems[2].Text, "", 0, 0, 0, 0, "NO (matching structure name)", tn, 0, gn));
                            }

                            if (matched > 1)  // multiple synonyms in absence of exact match
                            {
                                foreach (OLV_MappingDetails md in MappingDetailsOLVmodel)
                                {
                                    if (lvi.Text == md.templateName)
                                    {
                                        md.tnMatchCount = matched;
                                        md.group = "MULTIPLE SYNONYMS (for: '" + lvi.Text + "'; n = " + matched.ToString() + ")";
                                    }
                                }
                            }
                        }

                        mappingDetails_objectListView.SetObjects(MappingDetailsOLVmodel);
                    }
                }
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            TabPage current_tab = (sender as TabControl).SelectedTab;
            if (current_tab.Name == "valid_tabPage")
            {
                validTreeVsList_panel.Visible = true;
                if (Matches_listView.Visible)
                {
                    treeNavLinks_flowLayoutPanel.Visible = false;
                    currentPlanInTree_linkLabel.Visible = true;
                }
                else
                {
                    treeNavLinks_flowLayoutPanel.Visible = true;
                    collapseAll_linkLabel.Visible = true;
                    currentPlanInTree_linkLabel.Visible = false;
                }
            }
            else
            {
                validTreeVsList_panel.Visible = false;
                currentPlanInTree_linkLabel.Visible = false;
                treeNavLinks_flowLayoutPanel.Visible = true;
            }
        }

        private void validTree_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if ((sender as RadioButton).Checked)
            {
                Matches_listView.Visible = false;
                currentPlanInTree_linkLabel.Visible = false;
                treeNavLinks_flowLayoutPanel.Visible = true;
                exportValid_button.Visible = false;
            }
            else
            {
                Matches_listView.Visible = true;
                currentPlanInTree_linkLabel.Visible = true;
                treeNavLinks_flowLayoutPanel.Visible = false;
                exportValid_button.Visible = true;
            }
        }

        private void expandAll_linkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:  //valid plans
                    valid_treeView.BeginUpdate();
                    valid_treeView.ExpandAll();
                    valid_treeView.EndUpdate();
                    break;
                case 1:  //review
                    review_treeView.BeginUpdate();
                    review_treeView.ExpandAll();
                    review_treeView.EndUpdate();
                    break;
                case 2:  //none
                    none_treeView.BeginUpdate();
                    none_treeView.ExpandAll();
                    none_treeView.EndUpdate();
                    break;
                case 3:  //archived
                    archived_treeView.BeginUpdate();
                    archived_treeView.ExpandAll();
                    archived_treeView.EndUpdate();
                    break;
            }
        }

        private void collapseAll_linkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:  //valid plans
                    valid_treeView.CollapseAll();
                    break;
                case 1:  //review
                    review_treeView.CollapseAll();
                    break;
                case 2:  //none
                    none_treeView.CollapseAll();
                    break;
                case 3:  //archived
                    archived_treeView.CollapseAll();
                    break;
            }
        }

        private void currentPlanInTree_linkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Matches_listView.SelectedItems.Count == 1)
            {
                string mrn = Matches_listView.SelectedItems[0].SubItems[0].Text;
                string course_id = Matches_listView.SelectedItems[0].SubItems[1].Text;
                string plan_id = Matches_listView.SelectedItems[0].SubItems[2].Text;
                TreeNode structures_node = valid_treeView.Nodes[0].Nodes[mrn].Nodes[0].Nodes[course_id].Nodes[0].Nodes[plan_id].Nodes[0]; //no name set for levels with single node
                valid_treeView.CollapseAll();
                structures_node.Expand(); //show structures
                structures_node.EnsureVisible(); //expand parent nodes and scroll
                validTree_radioButton.Select(); //show tree via radio button as button visibilty linked to it
            }
        }


        //EXTRA FUNCTIONS (Begin)
        public void BoldThisNodeinThisTreeView(TreeNode node, TreeView tv)
        {
            node.NodeFont = new Font(tv.Font, FontStyle.Bold);
            node.Text += string.Empty;
        }
        public string GetKeyPath(TreeNode node)  //recursive function to get FullKEYPath of a tree node
        {
            if (node.Parent == null)
            {
                return node.Name;
            }

            return GetKeyPath(node.Parent) + " \\ " + node.Name;
        }

        void ExpandToLevel(TreeNodeCollection nodes, int level)
        {
            if (level > 0)
            {
                foreach (TreeNode node in nodes)
                {
                    node.Expand();
                    ExpandToLevel(node.Nodes, level - 1);
                }
            }
        }
        //EXTRA FUNCTIONS (End)

        private void button1_Click_1(object sender, EventArgs e)
        {
            // Login
            try
            {
                app = VMS.TPS.Common.Model.API.Application.CreateApplication(null, null);
                Console.WriteLine("Logged in as: " + app.CurrentUser.Id + Environment.NewLine);


                //EnableButton(SelectMRNList_Button);
                //EnableButton(DVHExport_Button);

            }
            catch (Exception exception)
            {
                MessageBox.Show("Error during login: " + exception.Message);
            }


            pullDosimetricData();

            //Logout
            if (app != null)
            {

                try
                {
                    app.Dispose();
                    Console.WriteLine(Environment.NewLine + "Logged out ");
                }
                catch (Exception exception)
                {
                    MessageBox.Show("Error during login: " + exception.Message);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            filterDosimetricData();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.DefaultExt = "json";
            dlg.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\raw"))
            {
                dlg.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\raw";
            }
            else { dlg.InitialDirectory = @"C:\temp\StructureNameAnalyser\data\raw"; }
            DialogResult result = dlg.ShowDialog();

            if (result == DialogResult.OK)
            {
                // serialize JSON directly to a file
                using (StreamWriter file = File.CreateText(dlg.FileName))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    serializer.Serialize(file, localPatients);
                }
            }


        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog = new System.Windows.Forms.OpenFileDialog();
            OpenFileDialog.DefaultExt = "json";
            OpenFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";

            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\raw"))
            {
                OpenFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\raw";
            }
            else { OpenFileDialog.InitialDirectory = @"C:\temp\StructureNameAnalyser\data\raw"; }


            try
            {

                if (OpenFileDialog.ShowDialog() == DialogResult.OK)
                {

                    // deserialize JSON directly from a file
                    using (StreamReader file = File.OpenText(OpenFileDialog.FileName))
                    {
                        JsonSerializer serializer = new JsonSerializer();
                        localPatients = (List<LocalPatient>)serializer.Deserialize(file, typeof(List<LocalPatient>));
                        indicatorLocalData_label.Text = localPatients.Count.ToString() + " loaded locally";
                        indicatorLocalData_label.Visible = true;
                        indicatorArchived_label.Visible = false;
                    }
                }
            }
            catch (Exception errorMsg)
            {
                MessageBox.Show(errorMsg.Message, "Error reading a file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void earliestCourseDate_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            earliestCourseDat_panel.Enabled = (sender as CheckBox).Checked;
            //sameNumberStructures_checkBox.Enabled = (sender as CheckBox).Checked;
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            sameNumberStructures_panel.Enabled = (sender as CheckBox).Checked;
            sameStructures_panel.Enabled = (sender as CheckBox).Checked;
        }

        private void sameNumberStructures_checkBox_EnabledChanged(object sender, EventArgs e)
        {
            //if (!sameNumberStructures_checkBox.Enabled) { sameNumberStructures_checkBox.Checked = false; }
        }

        private void expandCourses_linkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:  //valid plans
                    valid_treeView.BeginUpdate();
                    valid_treeView.CollapseAll();
                    ExpandToLevel(valid_treeView.Nodes, 3);
                    valid_treeView.Nodes[0].EnsureVisible();
                    valid_treeView.EndUpdate();
                    break;
                case 1:  //review
                    review_treeView.BeginUpdate();
                    review_treeView.CollapseAll();
                    ExpandToLevel(review_treeView.Nodes, 3);
                    review_treeView.Nodes[0].EnsureVisible();
                    review_treeView.EndUpdate();
                    break;
                case 2:  //none
                    none_treeView.BeginUpdate();
                    none_treeView.CollapseAll();
                    ExpandToLevel(none_treeView.Nodes, 3);
                    none_treeView.Nodes[0].EnsureVisible();
                    none_treeView.EndUpdate();
                    break;
                case 3:  //archived
                    archived_treeView.BeginUpdate();
                    archived_treeView.CollapseAll();
                    ExpandToLevel(archived_treeView.Nodes, 3);
                    archived_treeView.Nodes[0].EnsureVisible();
                    archived_treeView.EndUpdate();
                    break;
            }
        }

        private void expandPlans_linkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:  //valid plans
                    valid_treeView.BeginUpdate();
                    valid_treeView.CollapseAll();
                    ExpandToLevel(valid_treeView.Nodes, 5);
                    valid_treeView.Nodes[0].EnsureVisible();
                    valid_treeView.EndUpdate();
                    break;
                case 1:  //review
                    review_treeView.BeginUpdate();
                    review_treeView.CollapseAll();
                    ExpandToLevel(review_treeView.Nodes, 5);
                    review_treeView.Nodes[0].EnsureVisible();
                    review_treeView.EndUpdate();
                    break;
                case 2:  //none
                    none_treeView.BeginUpdate();
                    none_treeView.CollapseAll();
                    ExpandToLevel(none_treeView.Nodes, 5);
                    none_treeView.Nodes[0].EnsureVisible();
                    none_treeView.EndUpdate();
                    break;
                case 3:  //archived
                    archived_treeView.BeginUpdate();
                    archived_treeView.CollapseAll();
                    ExpandToLevel(archived_treeView.Nodes, 5);
                    archived_treeView.Nodes[0].EnsureVisible();
                    archived_treeView.EndUpdate();
                    break;
            }
        }

        private void linkLabel2_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:  //valid plans
                    valid_treeView.BeginUpdate();
                    valid_treeView.CollapseAll();
                    ExpandToLevel(valid_treeView.Nodes, 1);
                    valid_treeView.Nodes[0].EnsureVisible();
                    valid_treeView.EndUpdate();
                    break;
                case 1:  //review
                    review_treeView.BeginUpdate();
                    review_treeView.CollapseAll();
                    ExpandToLevel(review_treeView.Nodes, 1);
                    review_treeView.Nodes[0].EnsureVisible();
                    review_treeView.EndUpdate();
                    break;
                case 2:  //none
                    none_treeView.BeginUpdate();
                    none_treeView.CollapseAll();
                    ExpandToLevel(none_treeView.Nodes, 1);
                    none_treeView.Nodes[0].EnsureVisible();
                    none_treeView.EndUpdate();
                    break;
                case 3:  //archived
                    archived_treeView.BeginUpdate();
                    archived_treeView.CollapseAll();
                    ExpandToLevel(archived_treeView.Nodes, 1);
                    archived_treeView.Nodes[0].EnsureVisible();
                    archived_treeView.EndUpdate();
                    break;
            }
        }

        private void expandStructures_linkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            {
                switch (tabControl1.SelectedIndex)
                {
                    case 0:  //valid plans
                        valid_treeView.BeginUpdate();
                        valid_treeView.CollapseAll();
                        ExpandToLevel(valid_treeView.Nodes, 6);
                        valid_treeView.Nodes[0].EnsureVisible();
                        valid_treeView.EndUpdate();
                        break;
                    case 1:  //review
                        review_treeView.BeginUpdate();
                        review_treeView.CollapseAll();
                        ExpandToLevel(review_treeView.Nodes, 6);
                        review_treeView.Nodes[0].EnsureVisible();
                        review_treeView.EndUpdate();
                        break;
                    case 2:  //none
                        none_treeView.BeginUpdate();
                        none_treeView.CollapseAll();
                        ExpandToLevel(none_treeView.Nodes, 6);
                        none_treeView.Nodes[0].EnsureVisible();
                        none_treeView.EndUpdate();
                        break;
                    case 3:  //archived
                        archived_treeView.BeginUpdate();
                        archived_treeView.CollapseAll();
                        ExpandToLevel(archived_treeView.Nodes, 6);
                        archived_treeView.Nodes[0].EnsureVisible();
                        archived_treeView.EndUpdate();
                        break;
                }
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //read in static archived MRNs from same directory as exe
            string exe_path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (File.Exists(exe_path + "\\archivedMRNs.txt")){
                archivedMRNs = File.ReadAllLines(exe_path + "\\archivedMRNs.txt").ToList();
            }

                

            template_richTextBox.SelectionTabs = new int[] { 100, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800, 850, 900, 950, 1000 }; //set tabstop width (in pixels)
            faFilters.Add(new FaFilter("Example 1: containing 'PTV' or 'CTV'", "PTV|CTV"));
            faFilters.Add(new FaFilter("Example 2: containing 'GTV' but not 'MRI' or 'PET'", "^(?!.*(PET|MRI)).*GTV.*$"));
            faFilters.Add(new FaFilter("Example 3: containing 'PTV' or 'EVAL' but not 'BST' or 'BOOST'", "^(?!.*(BST|BOOST)).*(?=.*PTV)(?=.*EVAL).*$"));
            foreach (FaFilter f in faFilters)
            {
                faFilters_listView.Items.Add(f.name);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void gotoMrn1_button_Click(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedTab.Name)
            {
                case "valid_tabPage":
                    gotoMrnResults(tabControl1.SelectedTab, valid_treeView);
                    break;
                case "review_tabPage":
                    gotoMrnResults(tabControl1.SelectedTab, review_treeView);
                    break;
                case "none_tabPage":
                    gotoMrnResults(tabControl1.SelectedTab, none_treeView);
                    break;
                case "archived_tabPage":
                    gotoMrnResults(tabControl1.SelectedTab, archived_treeView);
                    break;
            }
        }

        private void gotoMrnResults(TabPage selectedTab, TreeView targetTree)
        {
            if (selectedTab.Name == "valid_tabPage")
            {
                ListViewItem found_item = Matches_listView.Items.Find(gotoMrn1_textBox.Text, false).FirstOrDefault();
                if (found_item != null)
                {
                    found_item.Selected = true;
                    found_item.Focused = true;
                    found_item.EnsureVisible();
                }
            }
            if (targetTree.Nodes.Count > 0)
            {
                TreeNode found_node = targetTree.Nodes[0].Nodes.Find(gotoMrn1_textBox.Text, false).FirstOrDefault();
                if (found_node != null)
                {
                    targetTree.CollapseAll();
                    targetTree.SelectedNode = found_node;
                    found_node.Nodes[0].Expand();
                    found_node.Nodes[0].EnsureVisible(); //expand parent nodes and scroll
                }
            }
        }

        private void gotoMrn2_button_Click(object sender, EventArgs e)
        {
            ListViewItem found_item = MRN_listView.Items.Find(gotoMrn2_textBox.Text, false).FirstOrDefault();
            if (found_item != null)
            {
                found_item.Selected = true;
                found_item.Focused = true;
                found_item.EnsureVisible();
            }
        }

        private void MRN_listView_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Determine whether the column is the same as the last column clicked.
            if (e.Column != mrnSortColumn)
            {
                // Set the sort column to the new column.
                mrnSortColumn = e.Column;
                // Set the sort order to ascending by default.
                MRN_listView.Sorting = SortOrder.Ascending;
            }
            else
            {
                // Determine what the last sort order was and change it.
                if (MRN_listView.Sorting == SortOrder.Ascending)
                    MRN_listView.Sorting = SortOrder.Descending;
                else
                    MRN_listView.Sorting = SortOrder.Ascending;

                // Call the sort method to manually sort.
                MRN_listView.Sort();
                // Set the ListViewItemSorter property to a new ListViewItemComparer
                // object.
                this.MRN_listView.ListViewItemSorter = new ListViewItemComparer(e.Column,
                                                                  MRN_listView.Sorting);
            }
        }

        class ListViewItemComparer : IComparer
        {
            private int col;
            private SortOrder order;
            public ListViewItemComparer()
            {
                col = 0;
                order = SortOrder.Ascending;
            }
            public ListViewItemComparer(int column, SortOrder order)
            {
                col = column;
                this.order = order;
            }
            [DebuggerStepThrough()]
            public int Compare(object x, object y)
            {
                int returnVal;
                // Determine whether the type being compared is a date type.
                try
                {
                    // Parse the two objects passed as a parameter as a DateTime.
                    System.DateTime firstDate =
                            DateTime.Parse(((ListViewItem)x).SubItems[col].Text);
                    System.DateTime secondDate =
                            DateTime.Parse(((ListViewItem)y).SubItems[col].Text);
                    // Compare the two dates.
                    returnVal = DateTime.Compare(firstDate, secondDate);
                }
                // If neither compared object has a valid date format, compare
                // as a string.
                catch
                {
                    // Compare the two items as a string.
                    returnVal = String.Compare(((ListViewItem)x).SubItems[col].Text,
                                ((ListViewItem)y).SubItems[col].Text);
                }
                // Determine whether the sort order is descending.
                if (order == SortOrder.Descending)
                    // Invert the value returned by String.Compare.
                    returnVal *= -1;
                return returnVal;
            }
        }

        private void Matrix_listView_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Determine whether the column is the same as the last column clicked.
            if (e.Column == 0 | e.Column == 3) //only make 'Date' column sortable
            {
                if (e.Column != matrixSortColumn)
                {
                    // Set the sort column to the new column.
                    matrixSortColumn = e.Column;
                    // Set the sort order to ascending by default.
                    Matrix_listView.Sorting = SortOrder.Ascending;
                }
                else
                {
                    // Determine what the last sort order was and change it.
                    if (Matrix_listView.Sorting == SortOrder.Ascending)
                        Matrix_listView.Sorting = SortOrder.Descending;
                    else
                        Matrix_listView.Sorting = SortOrder.Ascending;
                }
                // Call the sort method to manually sort.
                Matrix_listView.Sort();
                // Set the ListViewItemSorter property to a new ListViewItemComparer
                // object.
                this.Matrix_listView.ListViewItemSorter = new ListViewItemComparer(e.Column,
                                                                  Matrix_listView.Sorting);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {


        }

        private void button6_Click_1(object sender, EventArgs e)
        {

        }

        private void tumourSpecific_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            tumourSpecific_panel.Enabled = (sender as CheckBox).Checked;
            brCaSpecific_panel.Visible = (sender as CheckBox).Checked;

        }

        private void brCaLaterality_checkBox_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void brCaLatLTregex_textBox_TextChanged(object sender, EventArgs e)
        {

        }

        

        private void Matches_listView_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (e.Column != matchesSortColumn)
            {
                // Set the sort column to the new column.
                matchesSortColumn = e.Column;
                // Set the sort order to ascending by default.
                Matches_listView.Sorting = SortOrder.Ascending;
            }
            else
            {
                // Determine what the last sort order was and change it.
                if (Matches_listView.Sorting == SortOrder.Ascending)
                    Matches_listView.Sorting = SortOrder.Descending;
                else
                    Matches_listView.Sorting = SortOrder.Ascending;
            }
            // Call the sort method to manually sort.
            Matches_listView.Sort();
            // Set the ListViewItemSorter property to a new ListViewItemComparer
            // object.
            Matches_listView.ListViewItemSorter = new ListViewItemComparer(e.Column,
                                                              Matches_listView.Sorting);
        }

        private void AnalyseFrequCaseInsens_radioButton_radioButton_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void exportValid_button_Click(object sender, EventArgs e)
        {
            try
            {

                SaveFileDialog dlg = new SaveFileDialog();
                dlg.DefaultExt = "csv";
                dlg.Filter = "Tab delimited text (*.csv)|*.csv|All files (*.*)|*.*";

                if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\results"))
                {
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\data\results");
                }
                dlg.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\results";

                DialogResult result = dlg.ShowDialog();

                string fn;
                if (result == DialogResult.OK)
                {
                    // Save document
                    fn = dlg.FileName;
                    ListViewToCSV.DoListViewToCSV(Matches_listView, fn, true);
                    MessageBox.Show("Successfully filtered plans (or plan sums) exported as tab delimited text file to:" + Environment.NewLine + fn, "Export successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception errorMsg)
            {
                MessageBox.Show("Error during export! See detailed error message in brackets." + Environment.NewLine + "(" + errorMsg.Message + ")", "Export error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DelExclDose_button_Click(object sender, EventArgs e)
        {
            ExclDose_listBox.Items.Remove(ExclDose_listBox.SelectedItem);
        }

        private void AddExclDose_button_Click(object sender, EventArgs e)
        {
            //TODO: validation of input re allowed format via regex
            ExclDose_listBox.Items.Add(ExclDose_textBox.Text.Trim());
            ExclDose_textBox.Text = "";
        }

        private void AddInclDose_button_Click(object sender, EventArgs e)
        {
            //TODO: validation of input re allowed format via regex
            InclDose_listBox.Items.Add(InclDose_textBox.Text.Trim());
            InclDose_textBox.Text = "";
        }

        private void DelInclDose_button_Click(object sender, EventArgs e)
        {
            InclDose_listBox.Items.Remove(InclDose_listBox.SelectedItem);
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            try
            {

                SaveFileDialog dlg = new SaveFileDialog();
                dlg.DefaultExt = "xml";
                dlg.Filter = "XML file (*.xml)|*.xml|All files (*.*)|*.*";

                if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\filters"))
                {
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\data\filters");
                }
                dlg.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\filters";

                DialogResult result = dlg.ShowDialog();

                string fn;
                if (result == DialogResult.OK)
                {
                    // Save document
                    fn = dlg.FileName;
                    FormSerialisor.Serialise(this.splitContainer2.Panel1, fn);
                    MessageBox.Show("Filter settings successfully exported to:" + Environment.NewLine + fn, "Export successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception errorMsg)
            {
                MessageBox.Show("Error during export! See detailed error message in brackets." + Environment.NewLine + "(" + errorMsg.Message + ")", "Export error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            FormSerialisor.Deserialise(this.splitContainer2.Panel1, System.Windows.Forms.Application.StartupPath + @"\serialise.xml");
        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            try
            {
                System.Windows.Forms.OpenFileDialog OpenFileDialog = new System.Windows.Forms.OpenFileDialog();
                if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\filters"))
                {
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\data\filters");
                }
                OpenFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\filters";
                OpenFileDialog.Filter = "XML files (*.xml)|*.xml|All files (*.*)|*.*";
                OpenFileDialog.RestoreDirectory = true;

                if (OpenFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FormSerialisor.Deserialise(this.splitContainer2.Panel1, OpenFileDialog.FileName);
                }
            }
            catch (Exception errorMsg)
            {
                MessageBox.Show(errorMsg.Message, "Error reading a file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void mappingDetails_objectListView_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void matrixFullRowSelect_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (Matrix_listView.FullRowSelect)
            {
                Matrix_listView.FullRowSelect = false;
            }
            else
            {
                Matrix_listView.FullRowSelect = true;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            mappingDetails_objectListView.SelectAll();
            mappingDetails_objectListView.CopyObjectsToClipboard(mappingDetails_objectListView.SelectedObjects);
            mappingDetails_objectListView.SelectedObjects.Clear();
        }

        private void mappingDetails_objectListView_FormatCell(object sender, BrightIdeasSoftware.FormatCellEventArgs e)
        {
            if (e.ColumnIndex == this.mappingStatus_olvColumn.Index)
            {
                OLV_MappingDetails mps = (OLV_MappingDetails)e.Model;
                if (mps.mappingStatus.StartsWith("NAME"))
                {
                    e.SubItem.BackColor = Color.Green;
                }
                if (mps.mappingStatus.StartsWith("SYNONYM"))
                {
                    if (mps.tnMatchCount > 1)
                    {
                        e.SubItem.BackColor = Color.Yellow;
                    }
                    else
                    {
                        e.SubItem.BackColor = Color.Orange;
                    }
                    
                }
                if (mps.mappingStatus.StartsWith("NO (matching structure name)"))
                {
                    e.SubItem.BackColor = Color.Red;
                }
                if (mps.mappingStatus.StartsWith("EMPTY"))
                {
                    e.SubItem.BackColor = Color.LightBlue;
                }
            }
        }

        private void highlightFilter_button_Click(object sender, EventArgs e)
        {
            TextMatchFilter filter = TextMatchFilter.Contains(this.mappingDetails_objectListView, highlightFilter_textBox.Text);
            if (highlightFilter_checkBox.Checked)
            {
                this.mappingDetails_objectListView.ModelFilter = filter;
            }
            else
            {
                this.mappingDetails_objectListView.ModelFilter = null;
            }
            this.mappingDetails_objectListView.DefaultRenderer = new HighlightTextRenderer(filter);
        }

        private void highlightFilterClear_button_Click(object sender, EventArgs e)
        {
            this.mappingDetails_objectListView.ResetColumnFiltering();
            highlightFilter_textBox.Text = "";
        }

        private void Grouping_radioButtons_CheckedChanged(object sender, EventArgs e)
        {
            if (noGrouping_radioButton.Checked)
            {
                mappingDetails_objectListView.AlwaysGroupByColumn = null;
                mappingDetails_objectListView.Sort(index_olvColumn, SortOrder.Ascending);
                mappingDetails_objectListView.ShowGroups = false;
            }
            if (Grouping_radioButton.Checked)
            {
                mappingDetails_objectListView.AlwaysGroupByColumn = null;
                mappingDetails_objectListView.Sort(mappingStatus_olvColumn, SortOrder.Ascending);
                mappingDetails_objectListView.ShowGroups = true;
            }
            if (multipleSynonyms_radioButton.Checked)
            {
                mappingDetails_objectListView.ShowGroups = true;
                mappingDetails_objectListView.AlwaysGroupBySortOrder = SortOrder.Ascending;
                mappingDetails_objectListView.AlwaysGroupByColumn = group_olvColumn;
                mappingDetails_objectListView.Sort();
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            mappingDetails_objectListView.CopyObjectsToClipboard(mappingDetails_objectListView.SelectedObjects);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string csv = string.Empty;
            var olvExporter = new OLVExporter(mappingDetails_objectListView,
            mappingDetails_objectListView.FilteredObjects);
            csv = olvExporter.ExportTo(OLVExporter.ExportFormat.CSV);

            try
            {
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.DefaultExt = "csv";
                dlg.Filter = "Comma delimited text (*.csv)|*.csv|All files (*.*)|*.*";

                if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\results"))
                {
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\data\results");
                }
                dlg.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\results";

                DialogResult result = dlg.ShowDialog();

                if (result == DialogResult.OK)
                {
                    // Save document
                    using (StreamWriter sw = new StreamWriter(dlg.FileName))
                    {
                        sw.Write(csv);
                    }
                    //ListViewToCSV.DoListViewToCSV(Matches_listView, fn, true);
                    MessageBox.Show("Successfully exported as tab delimited text file to:" + Environment.NewLine + dlg.FileName, "Export successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception errorMsg)
            {
                MessageBox.Show("Error during export! See detailed error message in brackets." + Environment.NewLine + "(" + errorMsg.Message + ")", "Export error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void planRegex_textBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {
            //TODO: Should be refactored into separate routine used by this and and matrix on_click event handler

            MappingDetailsOLVmodel.Clear();

            mappingDetails_label2.Text = "ALL plans from matrix";

            foreach (ListViewItem Item in Matches_listView.Items)
            {
                List<StructureInfo> temp_structures_list = Item.Tag as List<StructureInfo>;
                if (temp_structures_list != null)
                {

                    string ms;
                    string tn;
                    int tnc;
                    string gn = "REMAINDER (matches exact or via single synonym and non-matches)";

                    //Go through all structures of current plan and attempt matching to template
                    foreach (StructureInfo st_iterator in temp_structures_list)
                    {
                        ms = "";
                        tn = "";
                        tnc = 0;


                        foreach (ListViewItem lvi in RepTemp_objectListView.Items)
                        {
                            Boolean leave_nested_loop = false;
                            string template_name = lvi.Text.Trim();
                            // List<string> synonyms = new List<string>();
                            Dictionary<string, string> synonyms_with_details = new Dictionary<string, string>();

                            if (lvi.SubItems.Count > 2)
                            {
                                //synonyms = lvi.SubItems[2].Text.Split(',').ToList();
                                List<string> temp_synonyms_with_details = new List<string>();
                                temp_synonyms_with_details = lvi.SubItems[2].Text.Split(',').ToList();
                                foreach (string syn_details in temp_synonyms_with_details)
                                {
                                    if (syn_details.Contains("{"))
                                    {
                                        string temp_syn = "";
                                        string temp_details = "";
                                        temp_syn = syn_details.Split('{')[0].Trim();
                                        temp_details = syn_details.Split('{')[1].Trim(new char[] { '}', ' ' });
                                        //Console.WriteLine("syn: '" + temp_syn + "'; " + "details: '" + temp_details + "'");
                                        synonyms_with_details.Add(temp_syn, temp_details);
                                    }
                                    else
                                    {
                                        synonyms_with_details.Add(syn_details, null);
                                    }
                                }
                            }

                            //structure name
                            if (template_name == st_iterator.Id)
                            {
                                if (st_iterator.isEmpty)
                                {
                                    ms = "EMPTY (name match but empty)";
                                }
                                else
                                {
                                    ms = "NAME (exact name match)";
                                }
                                tn = template_name;
                                break; //if name match don't look for synonyms by leaving foreach loop and goto next structure in plan
                            }
                            //check all synonyms of current template entry
                            for (int i = 0; i < synonyms_with_details.Count(); i++)
                            {
                                //if (synonyms[i].Trim() == st_iterator.Id)

                                if (synonyms_with_details.Keys.ElementAt(i).Trim() == st_iterator.Id)
                                {
                                    if (synonyms_with_details.Values.ElementAt(i) != null)  // with details condition
                                    {
                                        bool isConditionMet = false;
                                        if (synonyms_with_details.Values.ElementAt(i) != null) //check details expression via JINT (needs to evaluate to true or false)
                                        {
                                            var jintEval = new Engine()
                                                .SetValue("Vol", st_iterator.vol)
                                                .SetValue("Dmean", st_iterator.dMean)
                                                .SetValue("D95", st_iterator.d95)
                                                .Execute(synonyms_with_details.Values.ElementAt(i)) // details expression of current synonym
                                                .GetCompletionValue()
                                                .ToObject();
                                            isConditionMet = (bool)jintEval;
                                        }

                                        if (isConditionMet)
                                        {
                                            if (st_iterator.isEmpty)
                                            {
                                                ms = "EMPTY (syonym match but empty)";
                                            }
                                            else
                                            {
                                                ms = "SYNONYM (synonym match)";
                                            }
                                            tn = template_name;
                                            leave_nested_loop = true; //if synonym match don't look for further synonyms => leave nested loop 
                                            break;
                                        }
                                        else // synonym name match but not details => 
                                        {
                                            Console.WriteLine("All Details - MRN '" + Item.Text + "', Structure '" + st_iterator.Id + "' (Vol: "
                                             + st_iterator.vol.ToString() + "; Dmean: "
                                             + st_iterator.dMean.ToString() + ")  does not meet: '"
                                             + synonyms_with_details.Values.ElementAt(i) + "'");
                                        }
                                    }
                                    else  // no details condition
                                    {
                                        if (st_iterator.isEmpty)
                                        {
                                            ms = "EMPTY (syonym match but empty)";
                                        }
                                        else
                                        {
                                            ms = "SYNONYM (synonym match)";
                                        }
                                        tn = template_name;
                                        leave_nested_loop = true; //if synonym match don't look for further synonyms => leave nested loop 
                                        break;
                                    }
                                    
                                }
                                
                            }

                            if (leave_nested_loop)
                            {
                                break;
                            }

                        }

                        if (ms == "") //i.e. no exact or synonym match
                        {
                            ms = "NO (matching template name)";
                            tnc = 0;
                        }
                        else
                        {
                            tnc = 1; //if >1 will be fixed further down 
                        }

                        MappingDetailsOLVmodel.Add(new OLV_MappingDetails(MappingDetailsOLVmodel.Count + 1, Item.Text, Item.SubItems[1].Text, Item.SubItems[2].Text, st_iterator.Id, st_iterator.vol, st_iterator.d95, st_iterator.dMean, st_iterator.dMedian, ms, tn, tnc, gn));
                    }

                    //Determine template_names that couldn't be matched or where there are muliple matches (i.e. via multiple synonyms in absence of exact match)
                    foreach (ListViewItem lvi in RepTemp_objectListView.Items)
                    {
                        string temp_current_mrn = "";
                        int matched = 0;
                        foreach (OLV_MappingDetails md in MappingDetailsOLVmodel)
                        {
                            if (Item.Text == md.mrn)  //only look for current patient
                            {
                                temp_current_mrn = md.mrn;
                                if (lvi.Text == md.templateName)
                                {
                                    matched++;
                                }
                            }
                        }

                        if (Item.Text == temp_current_mrn)  //only look for current patient
                        {
                            if (matched == 0)  // no matches
                            {
                                tn = lvi.Text;
                                MappingDetailsOLVmodel.Add(new OLV_MappingDetails(MappingDetailsOLVmodel.Count + 1, Item.Text, Item.SubItems[1].Text, Item.SubItems[2].Text, "", 0, 0, 0, 0, "NO (matching structure name)", tn, 0, gn));
                            }

                            //if (matched > 1)  // multiple synonyms in absence of exact match
                            //{
                            //    foreach (OLV_MappingDetails md in MappingDetailsOLVmodel)
                            //    {
                            //        if (lvi.Text == md.templateName)
                            //        {
                            //            md.tnMatchCount = matched;
                            //            md.group = "MULTIPLE SYNONYMS (for: '" + lvi.Text + "'; n = " + matched.ToString() + ")";
                            //        }
                            //    }
                            //}
                        }
                    }

                    
                }
            }
            mappingDetails_objectListView.SetObjects(MappingDetailsOLVmodel);
        }

        private void detailsExportInclHeader_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            mappingDetails_objectListView.IncludeColumnHeadersInCopy = detailsExportInclHeader_checkBox.Checked;
        }

        private void splitContainer4_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                if (FAname_textBox.Text != "")
                {
                    FaFilter temp_FaFilter = new FaFilter(FAname_textBox.Text);
                    temp_FaFilter.frequencyTreshold = Convert.ToInt32(frequency_numericUpDown.Value);
                    temp_FaFilter.regEx = frequencyRegEx_textBox.Text;
                    temp_FaFilter.volMin = FAlowerVol_textBox.Text;
                    temp_FaFilter.volMax = FAupperVol_textBox.Text; 
                    temp_FaFilter.volTreshold = FAthreshVol_textBox.Text; 
                    if (dMean_radioButton.Checked){ temp_FaFilter.doseType = "Dmean"; } else { temp_FaFilter.doseType = "D95"; }
                    temp_FaFilter.doseMin = FAlowerDmean_textBox.Text; 
                    temp_FaFilter.doseMax = FAupperDmean_textBox.Text; 
                    temp_FaFilter.doseTreshold = FAthreshDmean_textBox.Text; 
                    faFilters.Add(temp_FaFilter);

                    faFilters_listView.Clear();
                    foreach (FaFilter f in faFilters)
                    {
                        faFilters_listView.Items.Add(f.name);
                    }
                }
            }
            catch { MessageBox.Show("Conversion error while trying to add Filter."); }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (faFilters_listView.SelectedIndices.Count > 0) { faFilters.RemoveAt(faFilters_listView.SelectedIndices[0]); }
            faFilters_listView.Clear();
            foreach (FaFilter f in faFilters)
            {
                faFilters_listView.Items.Add(f.name);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (faFilters_listView.SelectedIndices.Count > 0)
            {
                //get filter setting from global store
                FaFilter temp_FaFilter = faFilters[faFilters_listView.SelectedIndices[0]];

                //set in GUI
                FAname_textBox.Text = temp_FaFilter.name;
                frequency_numericUpDown.Value = temp_FaFilter.frequencyTreshold;
                frequencyRegEx_textBox.Text = temp_FaFilter.regEx;
                FAlowerVol_textBox.Text = temp_FaFilter.volMin;
                FAupperVol_textBox.Text = temp_FaFilter.volMax;
                FAthreshVol_textBox.Text = temp_FaFilter.volTreshold;
                if (temp_FaFilter.doseType == "Dmean") { dMean_radioButton.Checked = true; } else { D95_radioButton.Checked = true; }
                FAlowerDmean_textBox.Text = temp_FaFilter.doseMin;
                FAupperDmean_textBox.Text = temp_FaFilter.doseMax;
                FAthreshDmean_textBox.Text = temp_FaFilter.doseTreshold;
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog = new System.Windows.Forms.OpenFileDialog();
            OpenFileDialog.DefaultExt = "json";
            OpenFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";

            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\fa_filters"))
            {
                OpenFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\fa_filters";
            }
            else { OpenFileDialog.InitialDirectory = @"C:\temp\StructureNameAnalyser\data\fa_filters"; }


            try
            {

                if (OpenFileDialog.ShowDialog() == DialogResult.OK)
                {

                    // deserialize JSON directly from a file
                    using (StreamReader file = File.OpenText(OpenFileDialog.FileName))
                    {
                        JsonSerializer serializer = new JsonSerializer();
                        faFilters = (List<FaFilter>)serializer.Deserialize(file, typeof(List<FaFilter>));
                        faFilters_listView.Items.Clear();
                        foreach (FaFilter f in faFilters)
                        {
                            faFilters_listView.Items.Add(f.name);
                        }
                    }
                }
            }
            catch (Exception errorMsg)
            {
                MessageBox.Show(errorMsg.Message, "Error reading/loading frequency analysis filter file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.DefaultExt = "json";
            dlg.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\fa_filters"))
            {
                dlg.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\fa_filters";
            }
            else { dlg.InitialDirectory = @"C:\temp\StructureNameAnalyser\data\fa_filters"; }
            DialogResult result = dlg.ShowDialog();

            if (result == DialogResult.OK)
            {
                // serialize JSON directly to a file
                using (StreamWriter file = File.CreateText(dlg.FileName))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    serializer.Serialize(file, faFilters);
                }
            }
        }

        private void button6_Click_2(object sender, EventArgs e)
        {
            if (frequency_listView.CheckedItems.Count > 0)
            {
                {
                    if (frequency_listView.SelectedItems.Count == 0) //adding synonym to existing LSSN
                    {
                        if (template_objectListView.CheckedItems.Count == 1)
                        {
                            string lssn = template_objectListView.CheckedItems[0].Text;

                            foreach (ListViewItem lvi in frequency_listView.CheckedItems)
                            {
                                int rc = buildingTemplate.AddSynonymToClass(lssn, lvi.Text);
                                if (rc == 0)
                                {
                                    lvi.ForeColor = Color.Red;
                                    lvi.Checked = false;
                                }
                                else
                                {
                                    lvi.Checked = false;
                                    MessageBox.Show("Problem when adding synonym '" + lvi.Text + "'. Likely already exists in Template (Error code: " + rc + ").", "Problem while adding synonym");
                                }

                            }
                            template_objectListView.SetObjects(buildingTemplate.ContentAsListForOLV());
                            RepTemp_objectListView.SetObjects(buildingTemplate.ContentAsListForOLV());
                        }
                        else
                        {
                            MessageBox.Show("Can't associate synonym(s) with LSSN. Either highlight an entry in frequency list (creates new LSSN) or check *ONE* LSSN in dictonary (target LSSN).", "No LSSN selected", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else //adding new LSSN to dict
                    {
                        HashSet<string> synonyms = new HashSet<string>();
                        string lssn = "";
                        foreach (ListViewItem lvi in frequency_listView.CheckedItems)
                        {
                            if (lvi.Selected)
                            {
                                lssn = lvi.Text;
                            }
                            else
                            {
                                synonyms.Add(lvi.Text);
                            }
                            //lvi.ForeColor = SystemColors.InactiveCaption;
                            lvi.ForeColor = Color.Red;
                            lvi.Checked = false;
                        }
                        int rc = buildingTemplate.AddSemClass(new SemClass(lssn, groupName_textBox.Text, synonyms));
                        if (rc == 0)
                        {
                            template_objectListView.SetObjects(buildingTemplate.ContentAsListForOLV());
                            RepTemp_objectListView.SetObjects(buildingTemplate.ContentAsListForOLV());
                        } else
                        {
                            MessageBox.Show("Error code: " + rc, "ERROR while adding synonyms");
                        }
                    }

                    //unhighlight 
                    if (frequency_listView.SelectedItems.Count == 1) { frequency_listView.SelectedItems[0].Selected = false; }
                }
            }
        }

        private void LSSNdictV2_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (LSSNdictV2_radioButton.Checked)
            {
                LSSNdictV2_panel.Visible = true;
                LSSNdictV1_panel.Visible = false;
            }
            else
            {
                LSSNdictV1_panel.Visible = true;
                LSSNdictV2_panel.Visible = false;
            }
        }

        private void linkLabel1_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            foreach (ListViewItem lvi in frequency_listView.Items) { lvi.Selected = false; }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog = new OpenFileDialog();
            OpenFileDialog.DefaultExt = "txt";
            OpenFileDialog.Filter = "Tab delimited text (*.txt)|*.txt|All files (*.*)|*.*";

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\templates"))
            {
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\data\templates");
            }
            else { OpenFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\templates"; }

            try
            {
                if (OpenFileDialog.ShowDialog() == DialogResult.OK)
                {
                    buildingTemplate.LoadFromCSV(OpenFileDialog.FileName);
                    template_objectListView.SetObjects(buildingTemplate.ContentAsListForOLV());
                }
            }
            catch (Exception errorMsg)
            {
                MessageBox.Show(errorMsg.Message, "Error reading a file", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.DefaultExt = "txt";
                dlg.Filter = "Tab delimited text (*.txt)|*.txt|All files (*.*)|*.*";

                if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\templates"))
                {
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\data\templates");
                }
                dlg.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\templates";

                DialogResult result = dlg.ShowDialog();

                string fn;
                if (result == DialogResult.OK)
                {
                    // Save document
                    fn = dlg.FileName;
                    buildingTemplate.ContentToCSV(fn);
                    MessageBox.Show("Template exported as tab delimited text file to:" + Environment.NewLine + fn, "Export successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception errorMsg)
            {
                MessageBox.Show("Error during export! See detailed error message in brackets." + Environment.NewLine + "(" + errorMsg.Message + ")", "Export error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            Form2 modalDialog = new Form2();
            modalDialog.textBox1.Text = buildingTemplate.ContentAsString();
            if (modalDialog.ShowDialog() == DialogResult.OK)
            {
                // extract the data from the dialog
                int returnCode = buildingTemplate.LoadFromString(modalDialog.textBox1.Text);
                if (returnCode != 0)
                {
                    MessageBox.Show("Error code: " + returnCode, "ERROR while loading from string");

                }
                else
                {
                    template_objectListView.SetObjects(buildingTemplate.ContentAsListForOLV());
                    RepTemp_objectListView.SetObjects(buildingTemplate.ContentAsListForOLV());
                }
            }
        }

        private void treatedPlans_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            treatedPlan_panel.Enabled = (sender as CheckBox).Checked;
        }

        private void review_treeView_Click(object sender, EventArgs e)
        {

        }

        private void review_treeView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                Clipboard.SetText(e.Node.Text);
                review_treeView.SelectedNode = e.Node;
            }
        }

        private void none_treeView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                Clipboard.SetText(e.Node.Text);
                none_treeView.SelectedNode = e.Node;
            }
        }

        private void archived_treeView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                Clipboard.SetText(e.Node.Text);
                archived_treeView.SelectedNode = e.Node;
            }
        }

        private void template_objectListView_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            ////only allow one item to be checked at any time 
            foreach (OLVListItem item in template_objectListView.Items)
            {
                //uncheck all but to be checked
                if (!item.Equals(e.Item) && item.Checked)
                {
                    item.Checked = false;
                }
                //Unselect highlight in frequency listbox
                if (frequency_listView.SelectedItems.Count == 1) { frequency_listView.SelectedItems[0].Selected = false; }
            }
        }

        private void template_objectListView_ItemCheck(object sender, ItemCheckEventArgs e)
        {

        }

        private void template_objectListView_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void template_objectListView_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {

        }

        private void GotoStrucFrequ_button_Click(object sender, EventArgs e)
        {
            ListViewItem found_item = frequency_listView.Items.Find(GotoStrucFrequ_textBox.Text,false).FirstOrDefault();
            if (found_item != null)
            {
                found_item.Selected = true;
                found_item.Focused = true;
                found_item.EnsureVisible();
            }
        }

        private void HighlightDict_button_Click(object sender, EventArgs e)
        {
            TextMatchFilter filter = TextMatchFilter.Contains(this.template_objectListView, HighlightDict_textBox.Text);
            this.template_objectListView.ModelFilter = filter;
            this.template_objectListView.DefaultRenderer = new HighlightTextRenderer(filter);
        }

        private void HighlightDictClr_button_Click(object sender, EventArgs e)
        {
            this.template_objectListView.ResetColumnFiltering();
            HighlightDict_textBox.Text = "";
        }

        private void templLVheight_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (templLVheight_comboBox.SelectedIndex == 0) { template_objectListView.RowHeight = 16; }
            if (templLVheight_comboBox.SelectedIndex == 1) { template_objectListView.RowHeight = 32; }
            if (templLVheight_comboBox.SelectedIndex == 2) { template_objectListView.RowHeight = 48; }
            template_objectListView.Update();
        }

        private void template_objectListView_ItemCheck_1(object sender, ItemCheckEventArgs e)
        {
        }

        private void LoadRepTempFile_button_Click(object sender, EventArgs e)
        {
            OpenFileDialog OpenFileDialog = new OpenFileDialog();
            OpenFileDialog.DefaultExt = "txt";
            OpenFileDialog.Filter = "Tab delimited text (*.txt)|*.txt|All files (*.*)|*.*";

            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\data\templates"))
            {
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\data\templates");
            }
            else { OpenFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath + @"\data\templates"; }

            if (buildingTemplate.allMembersStem.Count > 0)
            {
                if (MessageBox.Show("There is a RISK you will override unsaved changes. ONLY proceed if you know that you have SAVED your previously built dictinary to file.", "Saved previous dictionary?", MessageBoxButtons.OKCancel,MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    try
                    {
                        if (OpenFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            buildingTemplate.LoadFromCSV(OpenFileDialog.FileName);
                            RepTemp_objectListView.SetObjects(buildingTemplate.ContentAsListForOLV());
                        }
                    }
                    catch (Exception errorMsg)
                    {
                        MessageBox.Show(errorMsg.Message, "Error reading a file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            } else
            {
                try
                {
                    if (OpenFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        buildingTemplate.LoadFromCSV(OpenFileDialog.FileName);
                        RepTemp_objectListView.SetObjects(buildingTemplate.ContentAsListForOLV());
                    }
                }
                catch (Exception errorMsg)
                {
                    MessageBox.Show(errorMsg.Message, "Error reading a file", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void perStrucV2_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (perStrucV2_radioButton.Checked)
            {
                LoadRepTempFile_button.Visible = true;
                RepTemp_objectListView.Visible = true;
            }else
            {
                LoadRepTempFile_button.Visible = false;
                RepTemp_objectListView.Visible = false;
            }
        }

        private void splitContainer5_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }

    public static class DateTimeExtensions
    {
        public static bool IsInRange(this DateTime dateToCheck, DateTime startDate, DateTime endDate)
        {
            return dateToCheck >= startDate && dateToCheck <= endDate;
            }
        }

        public static class ListViewToCSV
        {
            public static void DoListViewToCSV(ListView listView, string filePath, bool includeHidden)
            {
                //make header string
                StringBuilder result = new StringBuilder();
                WriteCSVRow(result, listView.Columns.Count, i => includeHidden || listView.Columns[i].Width > 0, i => listView.Columns[i].Text);

                //export data rows
                foreach (ListViewItem listItem in listView.Items)
                    WriteCSVRow(result, listItem.SubItems.Count, i => includeHidden || listView.Columns[i].Width > 0, i => listItem.SubItems[i].Text);

                File.WriteAllText(filePath, result.ToString());
            }

            private static void WriteCSVRow(StringBuilder result, int itemsCount, Func<int, bool> isColumnNeeded, Func<int, string> columnValue)
            {
                bool isFirstTime = true;
                for (int i = 0; i < itemsCount; i++)
                {
                    if (!isColumnNeeded(i))
                        continue;

                    if (!isFirstTime)
                        result.Append(",");
                    isFirstTime = false;

                    result.Append(String.Format("\"{0}\"", columnValue(i)));
                }
                result.AppendLine();
            }
        }


    }

