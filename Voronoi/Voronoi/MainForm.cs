using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;

using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Controls;
using ESRI.ArcGIS.ADF;
using ESRI.ArcGIS.SystemUI;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.AnalysisTools;
using System.Collections.Generic;

namespace Voronoi
{
    public sealed partial class MainForm : Form
    {
        #region 系统生成的方法
        #region class private members
        private IMapControl3 m_mapControl = null;
        private string m_mapDocumentName = string.Empty;
        #endregion

        #region class constructor
        public MainForm()
        {

            #region 初始化许可
            IAoInitialize m_AoInitialize = new AoInitializeClass();
            esriLicenseStatus licenseStatus = esriLicenseStatus.esriLicenseUnavailable;

            licenseStatus = m_AoInitialize.Initialize(esriLicenseProductCode.esriLicenseProductCodeArcInfo);
            //默认第一个为有效地，之后无效，此级别最高，可用绝大多数功能

            ////licenseStatus = m_AoInitialize.Initialize(esriLicenseProductCode.esriLicenseProductCodeEngine);级别最低

            #endregion

            InitializeComponent();
        }
        #endregion

        private void MainForm_Load(object sender, EventArgs e)
        {
            //get the MapControl
            m_mapControl = (IMapControl3)axMapControl1.Object;
        }

        #region Main Menu event handlers
        private void menuNewDoc_Click(object sender, EventArgs e)
        {
            //execute New Document command
            ICommand command = new CreateNewDocument();
            command.OnCreate(m_mapControl.Object);
            command.OnClick();
        }
        string fileName;
        string filePath;
        string strFullPath;
        IFeatureLayer pFeatureLayer;//点图层
        IFeatureLayer pFeatureLayer2;//生成的太森多边形图层
        private void menuOpenDoc_Click(object sender, EventArgs e)
        {
            ////execute Open Document command
            //ICommand command = new ControlsOpenDocCommandClass();
            //command.OnCreate(m_mapControl.Object);
            //command.OnClick();
            IWorkspaceFactory pWorkspaceFactory;
            IFeatureWorkspace pFeatureWorkspace;
            
            //获取当前路径和文件名
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "shpfile|*.shp";
            String ThiessenPolygoPath = @"C:\MyVoronoi\ThiessenPolygo.Shp";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                strFullPath = dlg.FileName;
                if (strFullPath == "") return;
                int Index = strFullPath.LastIndexOf("\\");
                filePath = strFullPath.Substring(0, Index);
                fileName = strFullPath.Substring(Index + 1);
                //打开工作空间并添加shp文件
                pWorkspaceFactory = new ShapefileWorkspaceFactoryClass();
                //注意此处的路径是不能带文件名的
                pFeatureWorkspace = (IFeatureWorkspace)pWorkspaceFactory.OpenFromFile(filePath, 0);
                pFeatureLayer = new FeatureLayerClass();
                //注意这里的文件名是不能带路径的
                pFeatureLayer.FeatureClass = pFeatureWorkspace.OpenFeatureClass(fileName);
                if (pFeatureLayer.FeatureClass.ShapeType != esriGeometryType.esriGeometryPoint)
                {
                    MessageBox.Show("shp文件不是点要素");
                    return;
                }
                pFeatureLayer.Name = pFeatureLayer.FeatureClass.AliasName;
                axMapControl1.Map.AddLayer(pFeatureLayer);
                axMapControl1.ActiveView.Refresh();
                menuOpenDoc.Enabled = false;

                if (!Directory.Exists("C:\\MyVoronoi"))
                {
                    Directory.CreateDirectory(@"C:\MyVoronoi");
                    //Directory.Delete("C:\\MyVoronoi", true);
                }
                //Directory.CreateDirectory(@"C:\MyVoronoi");

                CreateThiessenPolygons(pFeatureLayer, ThiessenPolygoPath, pFeatureWorkspace.ToString(), @"C:\MyVoronoi");

                //注意此处的路径是不能带文件名的
                pFeatureWorkspace = (IFeatureWorkspace)pWorkspaceFactory.OpenFromFile(@"C:\MyVoronoi", 0);
                pFeatureLayer2 = new FeatureLayerClass();
                //注意这里的文件名是不能带路径的
                pFeatureLayer2.FeatureClass = pFeatureWorkspace.OpenFeatureClass("ThiessenPolygo");
                pFeatureLayer2.Name = pFeatureLayer2.FeatureClass.AliasName;
                //axMapControl1.Map.AddLayer(pFeatureLayer2);
                //axMapControl1.ActiveView.Refresh();

                MessageBox.Show("初始化完毕！");
                axMapControl1.ContextMenuStrip = contextMenuStrip1;
                
            }

        }

        private void menuSaveDoc_Click(object sender, EventArgs e)
        {
            //execute Save Document command
            if (m_mapControl.CheckMxFile(m_mapDocumentName))
            {
                //create a new instance of a MapDocument
                IMapDocument mapDoc = new MapDocumentClass();
                mapDoc.Open(m_mapDocumentName, string.Empty);

                //Make sure that the MapDocument is not readonly
                if (mapDoc.get_IsReadOnly(m_mapDocumentName))
                {
                    MessageBox.Show("Map document is read only!");
                    mapDoc.Close();
                    return;
                }

                //Replace its contents with the current map
                mapDoc.ReplaceContents((IMxdContents)m_mapControl.Map);

                //save the MapDocument in order to persist it
                mapDoc.Save(mapDoc.UsesRelativePaths, false);

                //close the MapDocument
                mapDoc.Close();
            }
        }

        private void menuSaveAs_Click(object sender, EventArgs e)
        {
            //execute SaveAs Document command
            ICommand command = new ControlsSaveAsDocCommandClass();
            command.OnCreate(m_mapControl.Object);
            command.OnClick();
        }

        private void menuExitApp_Click(object sender, EventArgs e)
        {
            //exit the application
            Application.Exit();
        }
        #endregion

        //listen to MapReplaced evant in order to update the statusbar and the Save menu
        private void axMapControl1_OnMapReplaced(object sender, IMapControlEvents2_OnMapReplacedEvent e)
        {
            //get the current document name from the MapControl
            m_mapDocumentName = m_mapControl.DocumentFilename;

            //if there is no MapDocument, diable the Save menu and clear the statusbar
            if (m_mapDocumentName == string.Empty)
            {
                statusBarXY.Text = string.Empty;
            }
            else
            {
                //enable the Save manu and write the doc name to the statusbar
                statusBarXY.Text = System.IO.Path.GetFileName(m_mapDocumentName);
            }
        }

        private void axMapControl1_OnMouseMove(object sender, IMapControlEvents2_OnMouseMoveEvent e)
        {
            statusBarXY.Text = string.Format("{0}, {1}  {2}", e.mapX.ToString("#######.##"), e.mapY.ToString("#######.##"), axMapControl1.MapUnits.ToString().Substring(4));
        }
        #endregion
        #region MYMethod

        //调用gp工具创建泰森多边形
        public static void CreateThiessenPolygons(IFeatureLayer in_features, string out_feature_class, string workspacename, string extent)
        {
            //IAoInitialize m_AoInitialize = new AoInitializeClass();
            //esriLicenseStatus licenseStatus = esriLicenseStatus.esriLicenseUnavailable;
            //licenseStatus = m_AoInitialize.Initialize(esriLicenseProductCode.esriLicenseProductCodeArcInfo);
            ESRI.ArcGIS.Geoprocessor.Geoprocessor gp = new ESRI.ArcGIS.Geoprocessor.Geoprocessor();
            gp.OverwriteOutput = true;
            gp.SetEnvironmentValue("workspace", workspacename);
            //gp.SetEnvironmentValue("extent", extent);
            ESRI.ArcGIS.AnalysisTools.CreateThiessenPolygons createthiessenpolygon = new CreateThiessenPolygons();
            createthiessenpolygon.in_features = in_features;
            createthiessenpolygon.out_feature_class = out_feature_class;
            createthiessenpolygon.fields_to_copy = "ALL";
            gp.Execute(createthiessenpolygon, null);
            string strMessage = "";
            for (int i = 0; i < gp.MessageCount; i++)
            {
                strMessage += gp.GetMessage(i).ToString() + "\r\n";
            }
            //MessageBox.Show(strMessage);
        }

        private void tsmiKFG_Click(object sender, EventArgs e)
        {
            SelectScale frmSelectScale = new SelectScale();
            if (frmSelectScale.ShowDialog() == DialogResult.OK)
            {
                int OldScale = frmSelectScale.OldScale;
                int NewScale = frmSelectScale.NewScale;
                if (NewScale < OldScale)
                {
                    MessageBox.Show("无须删除！");
                    return;//无须筛选
                }
                int FeatureCound = pFeatureLayer.FeatureClass.FeatureCount(null);
                int NewCount = (int)(Math.Sqrt((double)OldScale / (double)NewScale) * (double)FeatureCound);
                int DelCount = FeatureCound - NewCount;//需要删除的数量

                List<s_P> listSp = new List<s_P>();

                ESRI.ArcGIS.Geodatabase.IQueryFilter queryFilter = new ESRI.ArcGIS.Geodatabase.QueryFilterClass();
                queryFilter.WhereClause = "1=1";
                IFeatureCursor featureCursor = pFeatureLayer2.Search(queryFilter, false);
                ESRI.ArcGIS.Geodatabase.IFeature pFeature;
                String PPP = "Input_FID";
                while ((pFeature = featureCursor.NextFeature()) != null)
                {
                    int index = pFeature.Fields.FindField(PPP);

                    if (pFeature.Shape is IArea)
                    {
                        s_P sp = new s_P();
                        sp.id = pFeature.get_Value(index).ToString();
                        sp.area = (pFeature.Shape as IArea).Area;
                        listSp.Add(sp);
                    }
                }

                //按照面积最小进行筛选
                List<s_P> listSpDel = new List<s_P>();
                for (int i = listSp.Count - 1; i >= 0; i--)
                {
                    if (listSpDel.Count == 0)
                    {
                        listSpDel.Add(listSp[i]);
                    }
                    else
                    {
                        for (int j=0;j< listSpDel.Count;j++)
                        {
                            if (listSpDel[j].area > listSp[i].area)
                            {
                                listSpDel.Insert(j + 1, listSp[i]);
                                break;
                            }
                        }
                    }
                }

                IDataset pDataset = pFeatureLayer.FeatureClass as IDataset;
                IWorkspace pWS = pDataset.Workspace;
                IWorkspaceEdit pWorkspaceEdit = pWS as IWorkspaceEdit;
                pWorkspaceEdit.StartEditing(false);
                pWorkspaceEdit.StartEditOperation();
                //删除点
                for(int j=listSpDel.Count-1;j>listSpDel.Count-1-DelCount; j--)
                {
                    ESRI.ArcGIS.Geodatabase.IQueryFilter queryFilter2 = new ESRI.ArcGIS.Geodatabase.QueryFilterClass();
                    queryFilter.WhereClause = String.Format("FID={0}", listSpDel[j].id);
                   
                    featureCursor = pFeatureLayer.Search(queryFilter2, false);
                    pFeature = featureCursor.NextFeature() ;
                    if (pFeature != null)
                    {
                        pFeature.Delete();
                    }
                }
                pWorkspaceEdit.StopEditOperation();
                pWorkspaceEdit.StopEditing(true);
                axMapControl1.Refresh();
                MessageBox.Show("处理完成,已经删除" + DelCount + "个点!");

            }
        }

        private void DeleteFeatures(IFeatureLayer pLayer, IQueryFilter queryFilter)  
        {  
          ITable pTable = pLayer.FeatureClass as ITable;  
          pTable.DeleteSearchedRows(queryFilter);  
        }  

        struct s_P
        {
            public String id;
            public double area;
        }

        private void menuFile_Click(object sender, EventArgs e)
        {

        }

        private void axMapControl1_OnMouseDown(object sender, IMapControlEvents2_OnMouseDownEvent e)
        {
            if (pFeatureLayer != null && pFeatureLayer2 != null)
            {
                contextMenuStrip1.Show(e.x + 191, e.y + 52);
            }
        }

        //public void CreateTin()
        //{
        //    ITin TinSurface;
        //    TinSurface =new Tin() as ITin;
        //    IFeatureClass FeatClass;
        //    IFeatureLayer pFeatureLayer;

        //    pFeatureLayer = axMapControl1.Map.get_Layer(0) as IFeatureLayer;
        //    IEnvelope pEnv;
        //    pEnv = pFeatureLayer.AreaOfInterest.Envelope;
        //    ITinEdit pTinEdit;
        //    pTinEdit = new Tin() as ITinEdit;
        //    pTinEdit.InitNew(pEnv);
        //    object o = new object();
        //    pTinEdit.SaveAs("MyTin", ref o);
        //    pTinEdit.StartEditing();
        //    FeatClass = pFeatureLayer.FeatureClass;
        //    IField pTagFeild;
        //    pTagFeild = new Field();
        //    IField pHightFeild;
        //    pHightFeild = FeatClass.Fields.get_Field(0);
        //    object oUseShapZ = new object(); 
        //    pTinEdit.AddFromFeatureClass(FeatClass, null, FeatClass.Fields.get_Field(0), FeatClass.Fields.get_Field(0), esriTinSurfaceType.esriTinMassPoint, ref oUseShapZ);
        //    pTinEdit.StopEditing(true);
        //    pTinEdit.Refresh();
        //    ITinNodeCollection pTinNodeCollection;
        //    pTinNodeCollection = pTinEdit as ITinNodeCollection;
        //    MessageBox.Show(pTinNodeCollection.NodeCount.ToString());
        //    ITin ptin;
        //    ptin = pTinEdit as ITin;
        //    IFeatureClass pNewfeatureclass;
        //    pNewfeatureclass = OpenFeatureClass_Example();
        //    pTinNodeCollection.ConvertToVoronoiRegions(pNewfeatureclass, null, null, "NodeIndex", "asdf");
        //}

        //public IFeatureClass OpenFeatureClass_Example()
        //{
        //    IWorkspaceFactory pWorkspaceFactory;
        //    pWorkspaceFactory = new ShapefileWorkspaceFactory();

        //    IFeatureWorkspace pFeatureWorkspace;
        //    pFeatureWorkspace = pWorkspaceFactory.OpenFromFile(Application.StartupPath + "\\MyTin", 0) as IFeatureWorkspace;
        //    IFeatureClass pFeatureClass;
        //    //elevation
        //    pFeatureClass = pFeatureWorkspace.OpenFeatureClass("*");
        //    pFeatureWorkspace.op
        //    return pFeatureClass;
        //}
        #endregion
    }
}