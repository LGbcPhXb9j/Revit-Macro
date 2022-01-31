/*
 * Created by SharpDevelop.
 * User: SKiiachko
 * Date: 5/12/2021
 * Time: 10:10 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Module
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("DD73EE28-9AA6-491E-975F-1909C9B86BAD")]
	public partial class ThisApplication
	{
		private void Module_Startup(object sender, EventArgs e)
		{

		}

		private void Module_Shutdown(object sender, EventArgs e)
		{

		}

		#region Revit Macros generated code
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Module_Startup);
			this.Shutdown += new System.EventHandler(Module_Shutdown);
		}
		#endregion
		public void Area()
		{

						Document doc = this.ActiveUIDocument.Document;
						                                         //create a filter to get all the title block type
                                         FilteredElementCollector FEC = new FilteredElementCollector(doc)
                                         .OfCategory(BuiltInCategory.OST_Rooms);
                                        TaskDialog td = new TaskDialog("Area report:");
                                        Transaction actrans = new Transaction(doc);
                                        actrans.Start("Room_Walls");
                                        string diatext=null;
                                        foreach (Room room in FEC){
                                        	SpatialElementGeometryCalculator calculator = new SpatialElementGeometryCalculator(doc);
											SpatialElementGeometryResults results = calculator.CalculateSpatialElementGeometry(room);
											Solid roomSolid = results.GetGeometry(); // get the solid representing the room's geometry
											    double sumArea=0;
										foreach (Face face in roomSolid.Faces)
{
    double faceArea = face.Area;

    IList<SpatialElementBoundarySubface> subfaceList = results.GetBoundaryFaceInfo(face); // get the sub-faces for the face of the room
    foreach (SpatialElementBoundarySubface subface in subfaceList)
    {
        if (subfaceList.Count > 1 && subface.SubfaceType == SubfaceType.Side) // there are multiple sub-faces that define the face
        {
            double subfaceArea = subface.GetSubface().Area; // get the area of each sub-face
            // sub-faces exist in situations such as when a room-bounding wall has been
            // horizontally split and the faces of each split wall combine to create the 
            // entire face of the room
            sumArea+=subfaceArea;
        }
    }
}
							sumArea*=0.092903;

							diatext +="\n Area of "+room.Name+": "+sumArea.ToString()+" sq.m";
                                        }
td.MainContent=diatext;
TaskDialogResult tdRes = td.Show();
                                                    actrans.Commit();
		}


		public void BOM()
		{
			double reserveP=1.0;
			double reserveD=1.0;
			bool quantity_override=false;
			Document doc = this.ActiveUIDocument.Document;
			FilteredElementCollector pipes = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_PipeCurves);
            FilteredElementCollector pfits = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_PipeFitting);
            var pfs= from FamilyInstance fi in pfits
            	where ((MechanicalFitting)fi.MEPModel).PartType==PartType.Elbow
            	select fi;
            FilteredElementCollector fpipes = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_FlexPipeCurves);
            FilteredElementCollector insul = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_PipeInsulations);
            FilteredElementCollector ducts = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_DuctCurves);
            FilteredElementCollector dfit = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_DuctFitting);
            FilteredElementCollector dinsul = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_DuctInsulations);
            FilteredElementCollector fducts = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_FlexDuctCurves);
            FilteredElementCollector types = new FilteredElementCollector(doc).OfClass(typeof(ElementType));
            Transaction ptrans = new Transaction(doc);
            ptrans.Start("BOM Data");
            List<List<Pipe>> listsp=new List<List<Pipe>>();
            List<List<Duct>> listsd=new List<List<Duct>>();
            List<List<FlexPipe>> listfp=new List<List<FlexPipe>>();
            List<List<FlexDuct>> listfd=new List<List<FlexDuct>>();
            List<List<PipeInsulation>> listspi=new List<List<PipeInsulation>>();
            List<List<DuctInsulation>> listsdi=new List<List<DuctInsulation>>();
            foreach (Pipe p in pipes)
            {
            	p.LookupParameter("STI_Length").Set(0);
            	List<List<Pipe>> tl=new List<List<Pipe>>(listsp);
            	if(listsp.Count==0) listsp.Add(new List<Pipe>(){p});
            	else
            	{
            		bool b=true;
            		foreach(List<Pipe> list in tl)
            		{
            			if(list[0].PipeType.Id==p.PipeType.Id&&list[0].Diameter==p.Diameter&&list[0].LookupParameter("STI_System").AsString()==p.LookupParameter("STI_System").AsString())
            			{
            				list.Add(p);
            				b=false;
            			}
            		}
            		if(b) listsp.Add(new List<Pipe>(){p});
            	}
            }
            
            foreach(List<Pipe> list in listsp)
            {
            	double fl=0;
            	foreach(FamilyInstance fi in pfs)
            	{
            		Connector con=null;
            		foreach(Connector c in fi.MEPModel.ConnectorManager.Connectors)
            		{
            			con=c;
            			break;
            		}
            		if(con.Radius*2==list[0].Diameter&&fi.Symbol.Name.IndexOf("bend", StringComparison.InvariantCultureIgnoreCase)==0&&fi.LookupParameter("STI_EN_Mark")==list[0].LookupParameter("STI_EN_Mark")&&list[0].LookupParameter("STI_System").AsString()==fi.LookupParameter("STI_System").AsString()&&fi.LookupParameter("STI_Length")!=null&&fi.LookupParameter("STI_Length").HasValue)
            		{
            			fl+=fi.LookupParameter("STI_Length").AsDouble();
            		
            		}
            }
            	double sum=list.Sum(x=>x.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble());
            	
            	double sum1=Math.Ceiling(UnitUtils.ConvertFromInternalUnits(reserveP*sum*2+fl,DisplayUnitType.DUT_METERS))-UnitUtils.ConvertFromInternalUnits(2*fl,DisplayUnitType.DUT_METERS);
            	double sum2=UnitUtils.ConvertFromInternalUnits(reserveP*sum*2,DisplayUnitType.DUT_METERS);
            	double factor=sum1/sum2;
            	foreach(Pipe pipe in list) pipe.LookupParameter("STI_Length").Set(pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()*factor);
            }
            foreach (FlexPipe f in fpipes)
            {
            	f.LookupParameter("STI_Length").Set(0);
            	List<List<FlexPipe>> tl=new List<List<FlexPipe>>(listfp);
            	if(listfp.Count==0) listfp.Add(new List<FlexPipe>(){f});
            	else
            	{
            		bool b=true;
            		foreach(List<FlexPipe> list in tl)
            		{
            			if(list[0].FlexPipeType.Id==f.FlexPipeType.Id&&list[0].Diameter==f.Diameter&&list[0].LookupParameter("STI_System").AsString()==f.LookupParameter("STI_System").AsString())
            			{
            				list.Add(f);
            				b=false;
            			}
            		}
            		if(b) listfp.Add(new List<FlexPipe>(){f});
            	}
            }
            foreach(List<FlexPipe> list in listfp)
            {
            	double sum=UnitUtils.ConvertFromInternalUnits(list.Sum(x=>x.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()),DisplayUnitType.DUT_METERS);
            	double factor=reserveP*Math.Ceiling(reserveP*2*sum)/(reserveP*2*sum);
            	foreach(FlexPipe pipe in list) pipe.LookupParameter("STI_Length").Set(pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()*factor);
            }
            foreach (PipeInsulation pi in insul)
            {
            	Element el=doc.GetElement(pi.HostElementId);
            	if(el!=null&&el.Category.Id.IntegerValue!=(int)BuiltInCategory.OST_PipeAccessory&&el.LookupParameter("STI_Length")!=null&&el.LookupParameter("STI_Length").HasValue)
            	{
            	pi.LookupParameter("STI_Length").Set(0);
            	List<List<PipeInsulation>> tl=new List<List<PipeInsulation>>(listspi);
            	if(listspi.Count==0) listspi.Add(new List<PipeInsulation>(){pi});
            	else
            	{
            		bool b=true;
            		foreach(List<PipeInsulation> list in tl)
            		{
            			Element inel=doc.GetElement(list[0].HostElementId);
            			double del=0,din=0;
            			if(el is Pipe)del=((Pipe)el).Diameter;
            			else
            			{
            				double rad=0;
            				foreach(Connector elc in ((FamilyInstance)el).MEPModel.ConnectorManager.Connectors)
            					if(elc.Radius>rad)rad=elc.Radius;
            				del=rad*2;
            			}
            			if(inel is Pipe)del=((Pipe)inel).Diameter;
            			else
            			{
            				double rad=0;
            				foreach(Connector elc in ((FamilyInstance)inel).MEPModel.ConnectorManager.Connectors)
            					if(elc.Radius>rad)rad=elc.Radius;
            				din=rad*2;
            			}
            			if(list[0].GetTypeId()==pi.GetTypeId()&&del==din&&list[0].Thickness==pi.Thickness&&list[0].LookupParameter("STI_System").AsString()==pi.LookupParameter("STI_System").AsString())
            			{
            				list.Add(pi);
            				b=false;
            			}
            		}
            		if(b) listspi.Add(new List<PipeInsulation>(){pi});
            	}
            	}
            }
            foreach(List<PipeInsulation> list in listspi)
            {
            	double sum=UnitUtils.ConvertFromInternalUnits(list.Sum(x=>doc.GetElement(x.HostElementId).LookupParameter("STI_Length").AsDouble()),DisplayUnitType.DUT_METERS);
            	double factor=Math.Ceiling(reserveP*2*sum)/(reserveP*2*sum);
            	foreach(PipeInsulation pinsul in list)pinsul.LookupParameter("STI_Length").Set(doc.GetElement(pinsul.HostElementId).LookupParameter("STI_Length").AsDouble()*factor);
            }
            foreach (Duct d in ducts)
            {
            	d.LookupParameter("STI_Length").Set(0);
            	List<List<Duct>> tl=new List<List<Duct>>(listsd);
            	if(listsd.Count==0) listsd.Add(new List<Duct>(){d});
            	else
            	{
            		bool b=true;
            		foreach(List<Duct> list in tl)
            		{
            			if((list[0].DuctType.Id==d.DuctType.Id&&list[0].LookupParameter("STI_System").AsString()==d.LookupParameter("STI_System").AsString())&&(((d.DuctType.Shape==ConnectorProfileType.Round)&&(list[0].Diameter==d.Diameter))||((d.DuctType.Shape==ConnectorProfileType.Rectangular)&&(list[0].Height==d.Height)&&(list[0].Width==d.Width))))
            			{
            				list.Add(d);
            				b=false;
            			}
            		}
            		if(b) listsd.Add(new List<Duct>(){d});
            	}
            }
            foreach(List<Duct> list in listsd)
            {
            	double sum=UnitUtils.ConvertFromInternalUnits(list.Sum(x=>x.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()),DisplayUnitType.DUT_METERS);
            	double factor=Math.Ceiling(reserveD*2*sum)/(reserveD*2*sum);
            	foreach(Duct duct in list) duct.LookupParameter("STI_Length").Set(duct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()*factor);
            }
            foreach (FlexDuct fd in fducts)
            {
            	fd.LookupParameter("STI_Length").Set(0);
            	List<List<FlexDuct>> tl=new List<List<FlexDuct>>(listfd);
            	if(listfd.Count==0) listfd.Add(new List<FlexDuct>(){fd});
            	else
            	{
            		bool b=true;
            		foreach(List<FlexDuct> list in tl)
            		{
            			if((list[0].FlexDuctType.Id==fd.FlexDuctType.Id&&list[0].LookupParameter("STI_System").AsString()==fd.LookupParameter("STI_System").AsString())&&(((fd.FlexDuctType.Shape==ConnectorProfileType.Round)&&(list[0].Diameter==fd.Diameter))||((fd.FlexDuctType.Shape==ConnectorProfileType.Rectangular)&&(list[0].Height==fd.Height)&&(list[0].Width==fd.Width))))
            			{
            				list.Add(fd);
            				b=false;
            			}
            		}
            		if(b) listfd.Add(new List<FlexDuct>(){fd});
            	}
            }
            foreach(List<FlexDuct> list in listfd)
            {
            	double sum=UnitUtils.ConvertFromInternalUnits(list.Sum(x=>x.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()),DisplayUnitType.DUT_METERS);
            	double factor=Math.Ceiling(reserveD*2*sum)/(reserveD*2*sum);
            	foreach(FlexDuct fduct in list) fduct.LookupParameter("STI_Length").Set(fduct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble()*factor);
            }
            foreach (DuctInsulation di in dinsul)
            {
            	Element el=doc.GetElement(di.HostElementId);
            	if((el.Category.Id.IntegerValue==(int)BuiltInCategory.OST_DuctCurves&&el.LookupParameter("STI_Length")!=null&&el.LookupParameter("STI_Length").HasValue)||(el.Category.Id.IntegerValue==(int)BuiltInCategory.OST_DuctFitting&&el.LookupParameter("!STI_Area")!=null&&el.LookupParameter("!STI_Area").HasValue))
            	{
            	di.LookupParameter("STI_Length").Set(0);
            	List<List<DuctInsulation>> tl=new List<List<DuctInsulation>>(listsdi);
            	if(listsdi.Count==0) listsdi.Add(new List<DuctInsulation>(){di});
            	else
            	{
            		bool b=true;
            		foreach(List<DuctInsulation> list in tl)
            		{
            			if(list[0].GetTypeId()==di.GetTypeId()&&list[0].Thickness==di.Thickness&&list[0].LookupParameter("STI_System").AsString()==di.LookupParameter("STI_System").AsString())
            			{
            				list.Add(di);
            				b=false;
            			}
            		}
            		if(b) listsdi.Add(new List<DuctInsulation>(){di});
            	}
            	}
            }
            foreach(List<DuctInsulation> list in listsdi)
            {
            	    double sum=0;
            		double sum2=0;
            	foreach(DuctInsulation duin in  list)
            	{
            		Element el=doc.GetElement(duin.HostElementId);
            		Duct duct=el as Duct;
            	if(el.Category.Id.IntegerValue==(int)BuiltInCategory.OST_DuctCurves)
            		{
            			double length=UnitUtils.ConvertFromInternalUnits(((Duct)el).get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(),DisplayUnitType.DUT_METERS);
            			if(duct.DuctType.Shape==ConnectorProfileType.Round)sum+=Math.PI*UnitUtils.ConvertFromInternalUnits((duct.Diameter),DisplayUnitType.DUT_METERS) * length;
                		else sum+=UnitUtils.ConvertFromInternalUnits(duct.Height+duct.Width,DisplayUnitType.DUT_METERS)*2*length;
            		}
            	            	
            	else
            	{

            		double x=UnitUtils.ConvertFromInternalUnits(el.LookupParameter("!STI_Area").AsDouble(),DisplayUnitType.DUT_SQUARE_METERS);
                			sum+=x;
                			sum2=x;
            	}
            	}
            	double factor=Math.Ceiling(reserveD*2*sum)/(2*sum);
            	foreach(DuctInsulation duin in  list)
            	{
            		Element el=doc.GetElement(duin.HostElementId);
            		Duct duct=el as Duct;
            	if(el.Category.Id.IntegerValue==(int)BuiltInCategory.OST_DuctCurves)
            		{
            			double length=UnitUtils.ConvertFromInternalUnits(((Duct)el).get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(),DisplayUnitType.DUT_METERS);
            			if(duct.DuctType.Shape==ConnectorProfileType.Round)duin.LookupParameter("STI_Length").Set(UnitUtils.ConvertToInternalUnits(factor*UnitUtils.ConvertFromInternalUnits((duct.Diameter),DisplayUnitType.DUT_METERS) * length*Math.PI, DisplayUnitType.DUT_METERS));
                		else duin.LookupParameter("STI_Length").Set(factor*UnitUtils.ConvertToInternalUnits(UnitUtils.ConvertFromInternalUnits(duct.Height+duct.Width,DisplayUnitType.DUT_METERS)*2*length, DisplayUnitType.DUT_METERS));
            		}
            	            	
            	else duin.LookupParameter("STI_Length").Set(UnitUtils.ConvertToInternalUnits(UnitUtils.ConvertFromInternalUnits(el.LookupParameter("!STI_Area").AsDouble()*factor,DisplayUnitType.DUT_SQUARE_METERS), DisplayUnitType.DUT_METERS));
            	}
            }
            foreach (FamilyInstance df in dfit)
            {
            	string str="";
            	Connector pr=null;
            	foreach(Connector con in df.MEPModel.ConnectorManager.Connectors)
            	{
            		if(pr==null||pr.Shape!=con.Shape||(pr.Shape==ConnectorProfileType.Round&&con.Shape==ConnectorProfileType.Round?pr.Radius!=con.Radius:pr.Height!=con.Height||pr.Width!=con.Width))
            		{
            			if(pr!=null)str+="-";
            			str+=con.Shape==ConnectorProfileType.Round?"Ø"+UnitUtils.ConvertFromInternalUnits(con.Radius*2,DisplayUnitType.DUT_MILLIMETERS).ToString():UnitUtils.ConvertFromInternalUnits(con.Height,DisplayUnitType.DUT_MILLIMETERS).ToString()+"x"+UnitUtils.ConvertFromInternalUnits(con.Width,DisplayUnitType.DUT_MILLIMETERS);
            		}
            		pr=con;
            	}
            df.LookupParameter("STI_Size").Set(str);
            }
            if(quantity_override)
            foreach(ElementType et in types)
            {
            	if(et.LookupParameter("STI_Quantity")!=null&&!et.LookupParameter("STI_Quantity").IsReadOnly)
            	   {
            	if(et is FamilySymbol)et.LookupParameter("STI_Quantity").Set(1);
            	else et.LookupParameter("STI_Quantity").Set(0);
            	}
            }
            ptrans.Commit();
        }
		public void FamilyParametersFixer()
		{
			string hz="";
		Document doc = this.ActiveUIDocument.Document;
		FilteredElementCollector fams = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(Family));
				Transaction t = new Transaction(doc);
		t.Start("Replace Family parameters");	
		FamilyManager fman=doc.FamilyManager;
			
		
			DefinitionFile spFile = this.Application.OpenSharedParameterFile();

						foreach (DefinitionGroup dG in spFile.Groups)
			{
						foreach (ExternalDefinition eD in dG.Definitions)
						{
							foreach (FamilyParameter fpam in fman.Parameters) 
							{
								ParameterType fpamt=fpam.Definition.ParameterType;
								ParameterType spamt=eD.ParameterType;
								if(eD.Name==fpam.Definition.Name&&!fpam.IsShared)
							{
								Dictionary<FamilyType,string> values=new Dictionary<FamilyType,string>();
								foreach (FamilyType ft in fman.Types) 
								{
									if (fpamt==ParameterType.Text)values.Add(ft,ft.AsString(fpam));
									else values.Add(ft,ft.AsValueString(fpam));
								}
								bool tb=fpam.IsInstance;
								BuiltInParameterGroup tbippg=fpam.Definition.ParameterGroup;
								fman.RemoveParameter(fpam);
								FamilyParameter newfp=fman.AddParameter(eD,tbippg,tb);
								foreach (FamilyType ft in fman.Types) 
								{

									fman.CurrentType=ft;
									fman.Set(newfp,values[ft]);
								}
								}

							}
						}
		}
		t.Commit();
		}
		public void RemoveUnusedParameters()
		{
			string parameters="";
					Document doc = this.ActiveUIDocument.Document;
		FilteredElementCollector elements = new FilteredElementCollector(doc).WhereElementIsNotElementType();
		FilteredElementCollector types = new FilteredElementCollector(doc).WhereElementIsElementType();
		FilteredElementCollector par = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(ParameterElement));
		List<Definition> usedparams=new List<Definition>();
		List<Definition> allparams=new List<Definition>();
//		        				Transaction t = new Transaction(doc);
//		        				 t.Start("Remove Empty Paarameters");	
//		        				foreach (ParameterElement p in par) 
//		        				{
//		        					if(p.Id.IntegerValue>0&&p.GetDefinition)
//		        					{
//		        					parameters+=p.Name+"\n";
//		        							        							        				TaskDialog.Show("ParameterElement",p.Name);
//		        					doc.Delete(p.Id);
//
//		        					}
//		        				}
//		        				            			t.Commit();

//            	foreach(Element el in elements)
//            	{
//            		foreach(Parameter p in el.Parameters)
//            		{
//            			p.Definition.
//            			if(p.Id.IntegerValue>0)
//            			{
//            				if(p.HasValue&&!usedparams.Contains(p.Definition))usedparams.Add(p.Definition);
//            			}
//            		}
//            	}
//            			t.Start("Add Space Shared Parameters");	
//            	foreach(Parameter par in sparam)
//            		if(!usedparams.Contains(par))
//            	{
//            		parameters+=par.Definition.Name+"\n";
//            		doc.ParameterBindings.Remove(par.Definition);
//            	}
//            			t.Commit();
		}
}
}
	
