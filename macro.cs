/*
 * Created by SharpDevelop.
 * User: SKiiachko
 * Date: 5/12/2021
 * Time: 10:10 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Text;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Drawing;
using System.Text.RegularExpressions;
using Group=Autodesk.Revit.DB.Group;
using Form=System.Windows.Forms.Form;
using TextBox=System.Windows.Forms.TextBox;
using Point=System.Drawing.Point;
using View=Autodesk.Revit.DB.View;
using Control=System.Windows.Forms.Control;


namespace Module
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("DD73EE28-9AA6-491E-975F-1909C9B86BAD")]
	public partial class ThisApplication
	{
		        static DialogResult ShowInputDialogBox(Dictionary<string,string> fields,string title = "Title",  int width = 300, int height = 150)
        {
            //This function creates the custom input dialog box by individually creating the different window elements and adding them to the dialog box

            //Specify the size of the window using the parameters passed

            //Create a new form using a System.Windows Form
            Form inputBox = new Form();

            inputBox.FormBorderStyle = FormBorderStyle.FixedDialog;

            //Set the window title using the parameter passed
            inputBox.Text = title;

            //Create a new label to hold the prompt
            int i=0;
            foreach (KeyValuePair<string,string> parameter in fields) {
            Label label = new Label();
            label.Text = parameter.Key;
            label.Location = new Point(5, i*50+5);
            label.Width = width - 15;
            inputBox.Controls.Add(label);

            //Create a textbox to accept the user's input
            TextBox textBox = new TextBox();
            textBox.Name=parameter.Key;
            textBox.Size = new Size(width - 10, 23);
            textBox.Location = new Point(5, label.Location.Y + 25);
            textBox.Text = parameter.Value;
            inputBox.Controls.Add(textBox);
            i++;
            }
            Size size = new Size(width, i*50+50);
            inputBox.ClientSize = size;

            //Create an OK Button 
            Button okButton = new Button();
            okButton.DialogResult = DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new Point(size.Width - 80 - 80, size.Height - 30);
            inputBox.Controls.Add(okButton);

            //Create a Cancel Button
            Button cancelButton = new Button();
            cancelButton.DialogResult = DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new Point(size.Width - 80, size.Height - 30);
            inputBox.Controls.Add(cancelButton);

            //Set the input box's buttons to the created OK and Cancel Buttons respectively so the window appropriately behaves with the button clicks
            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;

            //Show the window dialog box 
            DialogResult result = inputBox.ShowDialog();

            foreach (Control x in inputBox.Controls)
				{
				  if (x is TextBox)
				  {
				  	fields[((TextBox)x).Name]=((TextBox)x).Text ;
				  }
				}
            //After input has been submitted, return the input value
            return result;
        }

		 Dictionary<string,ParameterProperty> SharedParameters(List<ParameterProperty> names)
		{
		 	string message=string.Empty;
			Document doc=this.ActiveUIDocument.Document;
			Dictionary<string,ParameterProperty> guids=new Dictionary<string, ParameterProperty>();
			BindingMap map = doc.ParameterBindings;
			Dictionary<int,ParameterProperty> defs=new Dictionary<int,ParameterProperty>();
			DefinitionBindingMapIterator it= map.ForwardIterator();
			it.Reset();
				while(it.MoveNext())
						    {
				string spname=it.Key.Name;
				ParameterProperty pp=names.First(x=>x.Name==spname);
				if(pp!=null){
					defs.Add(((InternalDefinition)it.Key).Id.IntegerValue,pp);
					names.Remove(pp);
					if(names.Count==0)break;
										}
				}
			List<SharedParameterElement> spelist=new List<SharedParameterElement>(new FilteredElementCollector(doc).OfClass(typeof(SharedParameterElement)).Cast<SharedParameterElement>());
			foreach (SharedParameterElement spe in spelist) {
				int id=spe.GetDefinition().Id.IntegerValue;
				if (defs.ContainsKey(id) ){
					string n=spe.Name;
					ParameterProperty pp=defs[id];
					pp.GUID=spe.GuidValue;
					guids.Add(n,pp);
					defs.Remove(id);
					message+=spe.Name+"\n";
					}
				if(defs.Count==0){
					return guids;
				}
		}  
								return guids;			
		 }
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
			string discipline=doc.Title.Split('-')[1].Replace("3D",String.Empty);
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
            	if (discipline=="BH"){
            	double sum=UnitUtils.ConvertFromInternalUnits(list.Sum(x=>doc.GetElement(x.HostElementId).LookupParameter("STI_Length").AsDouble()),DisplayUnitType.DUT_METERS);
            	double factor=Math.Ceiling(reserveP*2*sum)/(reserveP*2*sum);
            	foreach(PipeInsulation pinsul in list)pinsul.LookupParameter("STI_Length").Set(doc.GetElement(pinsul.HostElementId).LookupParameter("STI_Length").AsDouble()*factor);
            	}
            	else
            {
            	double di=0;
            	double th=UnitUtils.ConvertFromInternalUnits(list[0].Thickness, DisplayUnitType.DUT_METERS);
            	string ss="";
            	if(doc.GetElement(list[0].HostElementId) is Pipe)
            	{
            			di=UnitUtils.ConvertFromInternalUnits(Math.PI*(((Pipe)doc.GetElement(list[0].HostElementId)).Diameter+list[0].Thickness), DisplayUnitType.DUT_METERS);
            			ss=di.ToString();
            	}
            	else
            	{
            		Connector c=null;
            		foreach (Connector con in ((FamilyInstance)doc.GetElement(list[0].HostElementId)).MEPModel.ConnectorManager.Connectors)
            			if(c==null||c.Radius<con.Radius)c=con;
            		
            		di=UnitUtils.ConvertFromInternalUnits(c.Radius*2, DisplayUnitType.DUT_METERS);
            		ss=di.ToString();
            	}
            	double sum=list.Sum(x=>UnitUtils.ConvertFromInternalUnits(doc.GetElement(x.HostElementId).LookupParameter("STI_Length").AsDouble(),DisplayUnitType.DUT_METERS));
            	double factor=Math.Ceiling(reserveP*di*th*200*sum)/(reserveP*di*th*200*sum);
            	foreach(PipeInsulation pinsul in list)pinsul.LookupParameter("STI_Length").Set(UnitUtils.ConvertToInternalUnits(UnitUtils.ConvertFromInternalUnits(doc.GetElement(pinsul.HostElementId).LookupParameter("STI_Length").AsDouble(), DisplayUnitType.DUT_METERS)*di*th, DisplayUnitType.DUT_METERS)*factor);
            }
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
            	if(el!=null)
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
																					string init="5 kg";
																					MatchCollection matches=Regex.Matches(init,@"\s+\d+(\.\d+)?$");
																					string one=(string)matches[0].Value;
										double two=0;
										double.TryParse(one,out two);
											TaskDialog.Show("Result",two.ToString());
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
		public void Test()
		{

			TaskDialog.Show(" hz", 100.ToString("500"));
		}
		public void TextReplacer()
		{
			Dictionary<string,string> fields=new Dictionary<string, string>(){
				{"Find", "3240"},
				{"Replace", "3140"}
			};
			ShowInputDialogBox(fields,"Replace input");
			string find=fields["Find"],replace=fields["Replace"];
			Document doc = this.ActiveUIDocument.Document;			
			ElementMulticategoryFilter mcf=new ElementMulticategoryFilter(new List<BuiltInCategory>(){
			                                                              BuiltInCategory.OST_ProjectInformation,
			                                                              BuiltInCategory.OST_Grids,
			                                                              BuiltInCategory.OST_Sheets,
			                                                              BuiltInCategory.OST_GenericAnnotation,
			                                                              BuiltInCategory.OST_Views,
			                                                              BuiltInCategory.OST_Schedules,
			                                                              BuiltInCategory.OST_Areas,
			                                                              BuiltInCategory.OST_PipeFitting,
			                                                              BuiltInCategory.OST_PipeCurves,
			                                                              BuiltInCategory.OST_GenericModel,
			                                                              BuiltInCategory.OST_MechanicalEquipment,
			                                                              BuiltInCategory.OST_PlumbingFixtures,
			                                                              BuiltInCategory.OST_PipeInsulations,
			                                                              BuiltInCategory.OST_PipeAccessory,
			                                                              BuiltInCategory.OST_SpecialityEquipment,
			                                                              BuiltInCategory.OST_DuctTerminal,
			                                                              BuiltInCategory.OST_DuctCurves,
			                                                              BuiltInCategory.OST_DuctAccessory,
			                                                              BuiltInCategory.OST_DuctFitting,
			                                                              BuiltInCategory.OST_DuctInsulations,
			                                                              BuiltInCategory.OST_Rooms,
			                                                              BuiltInCategory.OST_Walls,
			                                                              BuiltInCategory.OST_Floors,
			                                                              BuiltInCategory.OST_Roofs,
			                                                              BuiltInCategory.OST_Ceilings,
			                                                              BuiltInCategory.OST_Doors,
			                                                              BuiltInCategory.OST_Windows,
			                                                              BuiltInCategory.OST_Railings,
			                                                              BuiltInCategory.OST_StairsLandings,
			                                                              BuiltInCategory.OST_StairsRuns,
			                                                              BuiltInCategory.OST_Stairs,
			                                                              BuiltInCategory.OST_CurtainWallPanels,
			                                                              BuiltInCategory.OST_CurtainWallMullions,
			                                                              BuiltInCategory.OST_CableTrayFitting,
			                                                              BuiltInCategory.OST_CableTray,
			                                                              BuiltInCategory.OST_Conduit,
			                                                              BuiltInCategory.OST_ConduitFitting,
			                                                              BuiltInCategory.OST_ElectricalFixtures,
			                                                              BuiltInCategory.OST_ElectricalEquipment,
			                                                              BuiltInCategory.OST_LightingDevices,
			                                                              BuiltInCategory.OST_LightingFixtures,
			                                                              BuiltInCategory.OST_ElectricalCircuit
			                                                              });
			FilteredElementCollector elems = new FilteredElementCollector(doc).WherePasses(mcf);
			FilteredElementCollector tb = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_TextNotes);
			FilteredElementCollector views = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(View));
			FilteredElementCollector schedules = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(ViewSchedule));
			Transaction trans=new Transaction(doc);
			string s="";
			int number=0;
			trans.Start("Replace "+find+" with "+replace);
			foreach (View view in views) {
				if(view.Name.Contains(find))
				{
					view.Name=view.Name.Replace(find,replace);
					number++;
				}
			}
			foreach (ViewSchedule sch in schedules) {
				TableData td=sch.GetTableData();
					TableSectionData tsd=td.GetSectionData(SectionType.Header);
									for(int j=tsd.FirstColumnNumber;j<tsd.NumberOfColumns;j++)
				{
										if(tsd.IsValidColumnNumber(j))
														for(int k=tsd.FirstRowNumber;k<tsd.NumberOfRows;k++)
				{
											if(tsd.IsValidRowNumber(k)){}
										string t=tsd.GetCellText(k,j);
															if(t.Contains(find))
															{
																tsd.SetCellText(k,j,t.Replace(find,replace));
																number++;
															}
				}					
					
				}
			}
			foreach (Element el in elems) 
			{

				foreach (Parameter par in el.Parameters) {

											try{
						if(!par.IsReadOnly&&par.HasValue&&par.StorageType== StorageType.String&&par.AsString().Contains(find)){
						par.Set(par.AsString().Replace(find,replace));
					number++;
					}
											}
											catch(Exception e){s+="Error: "+el.Name+"|"+par.Definition.Name+"\n";}
				}

			}
			trans.Commit();
			s+=number.ToString()+" replaces";
						TaskDialog.Show("Note",s);
		}
		public void xXx()
		{
			Document doc = this.ActiveUIDocument.Document;			
			ElementMulticategoryFilter mcf=new ElementMulticategoryFilter(new List<BuiltInCategory>(){
			                                                              BuiltInCategory.OST_DuctTerminal,
			                                                              BuiltInCategory.OST_PipeFitting,
			                                                              BuiltInCategory.OST_PipeCurves,
			                                                              BuiltInCategory.OST_GenericModel,
			                                                              BuiltInCategory.OST_MechanicalEquipment,
			                                                              BuiltInCategory.OST_PlumbingFixtures,
			                                                              BuiltInCategory.OST_PipeInsulations,
			                                                              BuiltInCategory.OST_PipeAccessory,
			                                                              BuiltInCategory.OST_SpecialityEquipment,
			                                                              BuiltInCategory.OST_DuctCurves,
			                                                              BuiltInCategory.OST_DuctAccessory,
			                                                              BuiltInCategory.OST_DuctFitting,
			                                                              BuiltInCategory.OST_DuctInsulations,
			                                                              BuiltInCategory.OST_CableTrayFitting,
			                                                              BuiltInCategory.OST_CableTray,
			                                                              BuiltInCategory.OST_Conduit,
			                                                              BuiltInCategory.OST_ConduitFitting,
			                                                              BuiltInCategory.OST_ElectricalFixtures,
			                                                              BuiltInCategory.OST_ElectricalEquipment,
			                                                              BuiltInCategory.OST_LightingDevices,
			                                                              BuiltInCategory.OST_LightingFixtures,
			                                                              BuiltInCategory.OST_ElectricalCircuit
			                                                              });
			FilteredElementCollector elems = new FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(mcf);
			Transaction trans=new Transaction(doc);
			int number=0;
			string s="";
			trans.Start("xXx");
			foreach (Element el in elems) 
			{

				foreach (Parameter par in el.Parameters) {

											try{
						if(par.IsShared&&!par.IsReadOnly&&par.Definition.Name.StartsWith("AGC")&&(!par.HasValue||par.AsString()=="")){
						par.Set("XXXX");
					number++;
					}
											}
											catch(Exception e){s+="Error: "+el.Name+"|"+par.Definition.Name+"\n";}
				}

			}
			trans.Commit();
			s+=number.ToString()+" replaces";
						TaskDialog.Show("Note",s);
		}
		public void Ungroup()
		{
			Document doc = this.ActiveUIDocument.Document;			
			FilteredElementCollector elems = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(Group));
			Transaction trans=new Transaction(doc);
			int number=0;
			string s="";
			trans.Start("Ungroup");
				foreach (Group gr in elems) {
					gr.UngroupMembers();
					number++;
					
				}
			trans.Commit();
			s+=number.ToString()+" ungroups";
						TaskDialog.Show("Note",s);
		}
		public void Insulation()
		{
			string insulation="IC";
			Document doc = this.ActiveUIDocument.Document;			
			FilteredElementCollector elems = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_PipeInsulations);
			Transaction trans=new Transaction(doc);
			trans.Start("Ungroup");
				foreach (PipeInsulation pi in elems) {
				if(pi.HostElementId==null||doc.GetElement(pi.HostElementId).Category.Id.IntegerValue==(int)BuiltInCategory.OST_PipeAccessory)doc.Delete(pi.Id);
				else 
				{
					pi.LookupParameter("Insulation Code").Set(insulation);
					doc.GetElement(pi.HostElementId).LookupParameter("Insulation Code").Set(insulation);
					
				}
					
				}
			trans.Commit();
//			s+=number.ToString()+" ungroups";
//						TaskDialog.Show("Note",s);
		}
		public void ShowTag()
		{
			Document doc = this.ActiveUIDocument.Document;			
			FilteredElementCollector eq = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_MechanicalEquipment);
			FilteredElementCollector term = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_DuctTerminal);
			Transaction trans=new Transaction(doc);
			trans.Start("Show Tag");
				foreach (FamilyInstance equi in eq) {
				if(equi.LookupParameter("Show Tag")!=null)
				equi.LookupParameter("Show Tag").Set((int)1);
				}
							foreach (FamilyInstance termin in term) {
				if(termin.LookupParameter("Show Tag")!=null)
				termin.LookupParameter("Show Tag").Set((int)2);
				}
			trans.Commit();
		}
		public void TitleBlock()
		{
			Document doc = this.ActiveUIDocument.Document;
			Dictionary<string,string> input=new Dictionary<string, string>(){
				{"Project Name","AGCC"},
				{"Client Name","TCM"},
				{"RF Code Control","O. Tokar"},
				{"PM","M. Perelli"},
			};
			Dictionary<string,string> runames=new Dictionary<string, string>(){
					{"A.Alexandrov","A. Александров"},
					{"A.Borisova","A. Борисова"},
					{"A.Dolgikh","A. Долгих"},
					{"A.Danilchenko","A. Данильченко"},
					{"A.Kirillova","A. Кириллова"},
					{"A.Ozhiganov","A. Ожиганов"},
					{"A.Tsoi","A. Цой"},
					{"B.Berzhakanova","Б. Бержаканова"},
					{"D.Pichugin","Д. Пичугин"},
					{"F.Pak","Ф. Пак"},
					{"M.Abdrafikov","М. Абдрафиков"},
					{"M.Cargnelutti","М. Карнелутти"},
					{"M.Perelli","М. Перелли"},
					{"N.Barantseva","Н. Баранцева"},
					{"O.Tokar","О. Токарь"},
					{"P.Tokarev","П. Токарев"},
					{"S.Kostromin","С. Костромин"},
					{"V.Klunok","В. Клунок"},
					{"V.Guseva","С. Гусева"},
					{"V.Shutarev","В. Шутарев"},
					{"O.Mashkantsev","О. Машканцев"}
				};
			ShowInputDialogBox(input,"Common Parameters");
			List<string> parList=new List<string>(){"Discipline Code","WBS Number","Building Symbol","WBS Sequence Number","Marka Code","Construction Area"};
			ProjectInfo pinf=doc.ProjectInformation;
			Dictionary<String,Parameter> parameters=new Dictionary<string, Parameter>();
			Transaction trans=new Transaction(doc);
			trans.Start("Title Block");
			foreach (Parameter par in pinf.Parameters) {
				string name=par.Definition.Name;
				if(parList.Count>0&&parList.Contains(name)){
					parameters.Add(name,par);
					parList.Remove(name);}
				else if(par.Definition is InternalDefinition){
					BuiltInParameter bip=((InternalDefinition)par.Definition).BuiltInParameter;
					if(bip== BuiltInParameter.PROJECT_NAME)par.Set(input["Project Name"]);
					if(bip== BuiltInParameter.PROJECT_NUMBER)par.Set(doc.Title.Substring(0,doc.Title.IndexOf("-")));
					if(bip== BuiltInParameter.CLIENT_NAME)par.Set(input["Client Name"]);
					if(bip== BuiltInParameter.PROJECT_AUTHOR)par.Set(runames[input["PM"].Replace(" ",String.Empty)]);
					if(par.Definition.Name=="#STI_Control")par.Set(runames[input["RF Code Control"].Replace(" ",String.Empty)]);
				}
				}
			string discipline=doc.Title.Substring(doc.Title.IndexOf("-")+1,2);
		parameters["Discipline Code"].Set(discipline);
		parameters.Remove("Discipline Code");
			ElementMulticategoryFilter mcf=new ElementMulticategoryFilter(new List<BuiltInCategory>(){
			                                                              BuiltInCategory.OST_PipeFitting,
			                                                              BuiltInCategory.OST_PipeCurves,
			                                                              BuiltInCategory.OST_GenericModel,
			                                                              BuiltInCategory.OST_MechanicalEquipment,
			                                                              BuiltInCategory.OST_PlumbingFixtures,
			                                                              BuiltInCategory.OST_PipeInsulations,
			                                                              BuiltInCategory.OST_PipeAccessory,
			                                                              BuiltInCategory.OST_SpecialityEquipment,
			                                                              BuiltInCategory.OST_DuctTerminal,
			                                                              BuiltInCategory.OST_DuctCurves,
			                                                              BuiltInCategory.OST_DuctAccessory,
			                                                              BuiltInCategory.OST_DuctFitting,
			                                                              BuiltInCategory.OST_DuctInsulations,
			                                                              BuiltInCategory.OST_Rooms,
			                                                              BuiltInCategory.OST_Walls,
			                                                              BuiltInCategory.OST_Floors,
			                                                              BuiltInCategory.OST_Roofs,
			                                                              BuiltInCategory.OST_Ceilings,
			                                                              BuiltInCategory.OST_Doors,
			                                                              BuiltInCategory.OST_Windows,
			                                                              BuiltInCategory.OST_Railings,
			                                                              BuiltInCategory.OST_StairsLandings,
			                                                              BuiltInCategory.OST_StairsRuns,
			                                                              BuiltInCategory.OST_Stairs,
			                                                              BuiltInCategory.OST_CableTrayFitting,
			                                                              BuiltInCategory.OST_CableTray,
			                                                              BuiltInCategory.OST_Conduit,
			                                                              BuiltInCategory.OST_ConduitFitting,
			                                                              BuiltInCategory.OST_ElectricalFixtures,
			                                                              BuiltInCategory.OST_ElectricalEquipment,
			                                                              BuiltInCategory.OST_LightingDevices,
			                                                              BuiltInCategory.OST_LightingFixtures,
			                                                              BuiltInCategory.OST_ElectricalCircuit
			                                                              });
			FilteredElementCollector elems = new FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(mcf);
			if(discipline!="BB")
			foreach (Element el in elems) {
			foreach (Parameter par in el.Parameters) {
					if(!par.IsReadOnly&&par.IsShared&&par.HasValue&&par.StorageType==StorageType.String&&par.AsString()!=""&&parameters.ContainsKey(par.Definition.Name)){
						if(par.Definition.Name=="Construction Area")parameters[par.Definition.Name].Set(par.AsString().Insert(2,"-"));
						else parameters[par.Definition.Name].Set(par.AsString());
						parameters.Remove(par.Definition.Name);
						
					}
					if(parameters.Count==0)break;
			}	
			}
			else
			{
				Dictionary<string,string> arp=new Dictionary<string, string>(){
					{"LV00_WBS Area Code","WBS Number"},
					{"LV01_Object Code","Building Symbol"},
					{"LV01_Sequence Number","WBS Sequence Number"},
					{"PTY_04","Marka Code"},
					{"CWA","Construction Area"}
				};
							foreach (Element el in elems) {
			foreach (Parameter par in el.Parameters) {
					if(!par.IsReadOnly&&par.IsShared&&par.HasValue&&par.StorageType==StorageType.String&&par.AsString()!=""&&arp.ContainsKey(par.Definition.Name)){
					if(par.Definition.Name=="CWA")parameters[arp[par.Definition.Name]].Set(par.AsString().Insert(2,"-"));
					else parameters[arp[par.Definition.Name]].Set(par.AsString());
						arp.Remove(par.Definition.Name);
						
					}
					if(parameters.Count==0)break;
			}	
			}
			}
			FilteredElementCollector tb = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_TitleBlocks);
				foreach (FamilyInstance fi in tb) {
				if(fi.OwnerViewId!=null&&fi.Symbol.FamilyName=="STI_Title_Block"){
				ViewSheet vs=doc.GetElement(fi.OwnerViewId) as ViewSheet;
				int count=vs.GetAllRevisionIds().Count;
				fi.LookupParameter("Revisions").Set(count);
//				fi.LookupParameter("Changes").Set(count-vs.GetAdditionalRevisionIds().Count);
				string scale="N.T.S.";
					Viewport view=null;
					ElementCategoryFilter ecf=new ElementCategoryFilter(BuiltInCategory.OST_RevisionClouds);
					Dictionary<int,int> rcquan=new Dictionary<int, int>();
				foreach (ElementId id in vs.GetAllViewports()) {
						Viewport vp=doc.GetElement(id) as Viewport;
						List<ElementId> rcs=(List<ElementId>)doc.GetElement(vp.ViewId).GetDependentElements(ecf);
						if (rcs!=null)
						foreach (ElementId irc in rcs) {
							int rc=((RevisionCloud)doc.GetElement(irc)).RevisionId.IntegerValue;
							if(rcquan.ContainsKey(rc))rcquan[rc]++;
							else rcquan.Add(rc,1);
						}
						if(doc.GetElement(vp.ViewId) is ViewPlan)
						{
							if (view==null)view=vp;
							else if(Math.Sqrt(Math.Pow(vp.GetBoxOutline().MaximumPoint.X-vp.GetBoxOutline().MinimumPoint.X,2)+Math.Pow(vp.GetBoxOutline().MaximumPoint.Y-vp.GetBoxOutline().MinimumPoint.Y,2))>Math.Sqrt(Math.Pow(view.GetBoxOutline().MaximumPoint.X-view.GetBoxOutline().MinimumPoint.X,2)+Math.Pow(view.GetBoxOutline().MaximumPoint.Y-view.GetBoxOutline().MinimumPoint.Y,2)))view=vp;
						}
				}
						List<ElementId>vrcs=(List<ElementId>)vs.GetDependentElements(ecf);
						if (vrcs!=null)
						foreach (ElementId irc in vrcs) {
							int rc=((RevisionCloud)doc.GetElement(irc)).RevisionId.IntegerValue;
							if(rcquan.ContainsKey(rc))rcquan[rc]++;
							else rcquan.Add(rc,1);
						}
					if(view!=null)scale="1:"+((View)doc.GetElement(view.ViewId)).Scale.ToString();
				vs.LookupParameter("!STI_Scale").Set(scale);
				double number=vs.LookupParameter("!STI_D_Number").AsDouble();
				vs.LookupParameter("!STI_Number").Set(Math.Floor(number).ToString("0000"));
				if(vs.LookupParameter("!STI_Type").AsString()=="BOM"){
					vs.LookupParameter("!STI_Number").Set("0001");
					fi.LookupParameter("Form").Set(number>1?6:3);
				}
				else fi.LookupParameter("Form").Set(number%1>0.1000001?6:3);
				int i=0;
				Dictionary<int,string> values=new Dictionary<int, string>();
				foreach (ElementId revid in vs.GetAllRevisionIds().Reverse()) {
					string parname="#STI_SNPM";
					Revision rev=doc.GetElement(revid) as Revision;
					string[] names=rev.IssuedBy.Split(',');
					if (names.Length<3)break;
					try{
					values=new Dictionary<int, string>(){{1,rev.RevisionNumber.Contains('.')?rev.RevisionNumber.Substring(0,rev.RevisionNumber.IndexOf('.')):rev.RevisionNumber},{2,rev.RevisionDate},{3,rev.Description},{4,names[0]},{5,names[1]},{6,names[2]},{7,rcquan.ContainsKey(rev.Id.IntegerValue)?rcquan[rev.Id.IntegerValue].ToString():String.Empty}};

					if(i==0){

				vs.get_Parameter(BuiltInParameter.SHEET_DRAWN_BY).Set(runames[values[4].Replace(" ",String.Empty)]);
				vs.get_Parameter(BuiltInParameter.SHEET_CHECKED_BY).Set(runames[values[5].Replace(" ",String.Empty)]);
				vs.get_Parameter(BuiltInParameter.SHEET_APPROVED_BY).Set(runames[values[6].Replace(" ",String.Empty)]);
				ISet<ElementId> set=fi.Symbol.Family.GetFamilyTypeParameterValues(fi.LookupParameter("1_Sign_Drawer").Id);
				Dictionary<string,ElementId> types=set.ToDictionary(x=>doc.GetElement(x).Name);
				fi.LookupParameter("1_Sign_Drawer").Set(types[values[4].Replace(" ",String.Empty).Replace(".",String.Empty)]);
				fi.LookupParameter("2_Sign_Checker").Set(types[values[5].Replace(" ",String.Empty).Replace(".",String.Empty)]);
				fi.LookupParameter("3_Sign_Approver").Set(types[values[6].Replace(" ",String.Empty).Replace(".",String.Empty)]);
				fi.LookupParameter("4_Sign_Control").Set(types[input["RF Code Control"].Replace(".",String.Empty).Replace(" ",String.Empty)]);
				fi.LookupParameter("5_Sign_Manager").Set(types[input["PM"].Replace(".",String.Empty).Replace(" ",String.Empty)]);
					}
										}
					catch{
						string s="Incorrect names in revision or missing in database: \n";
						foreach (KeyValuePair<string, string> kvp in runames) {
							s+="["+kvp.Key+"]\n";
						}
						TaskDialog.Show("Error",s);
							break;
					}
					for (int j = 1; j < 8; j++) {
						Parameter par=vs.LookupParameter(parname.Replace("N",i.ToString()).Replace("M",j.ToString()));
						if(par!=null)par.Set(values[j]);
						
					}
					if(i==count){
					}
					i++;
				}

				i=0;
				}
			}

			trans.Commit();
		}
		enum Behavior{
			Common,
			Sequence,
			Unique,
			Combine,
			CombineCommon,
			Auto,
			None
		}
		class ParameterProperty{
			public string Name{get;set;}
			public string Discipline{get;set;}
			public string Argument{get;set;}
			public Guid GUID{get;set;}
			public Behavior Behavior{get;set;}
			public Dictionary<string,int> Counter=new Dictionary<string, int>();
			public Dictionary<string,string> Selector=new Dictionary<string, string>();
			public string Value(string target=""){
				switch (Behavior) {
						case Behavior.Common:return Selector.First().Value;
						
						break;
					default: return "";
						
						break;
				}
				}
			public ParameterProperty(string name,string discipline="",Behavior behavior= Behavior.None,Dictionary<string,string> arguments= null){
				Name=name;
				Discipline=discipline;
//				Argument=argument;
				Behavior= behavior;
			}
		}
		Dictionary<BuiltInCategory,string> Categories=new Dictionary<BuiltInCategory, string>(){
			{BuiltInCategory.OST_PipeFitting,"BU,BH,BB"},
			{BuiltInCategory.OST_PipeCurves,"BU,BH,BB"},
			{BuiltInCategory.OST_GenericModel,"BB,BN,BU,BJ,BK,BH"},
			{BuiltInCategory.OST_MechanicalEquipment,"BB,BU,BJ,BK,BH"},
			{BuiltInCategory.OST_PlumbingFixtures,"BB,BU,BH"},
			{BuiltInCategory.OST_PipeInsulations,"BB,BU,BH"},
			{BuiltInCategory.OST_PipeAccessory,"BB,BU,BH"},
			{BuiltInCategory.OST_SpecialityEquipment,"BB,BU,BH"},
			{BuiltInCategory.OST_DuctTerminal,"BH"},
			{BuiltInCategory.OST_DuctCurves,"BH,BN"},
			{BuiltInCategory.OST_DuctAccessory,"BH,BN"},
			{BuiltInCategory.OST_DuctFitting,"BH,BN"},
			{BuiltInCategory.OST_DuctInsulations,"BH,BN"},
			{BuiltInCategory.OST_Rooms,"BB"},
			{BuiltInCategory.OST_Walls,"BB"},
			{BuiltInCategory.OST_Floors,"BB"},
			{BuiltInCategory.OST_Roofs,"BB"},
			{BuiltInCategory.OST_Ceilings,"BB"},
			{BuiltInCategory.OST_Doors,"BB"},
			{BuiltInCategory.OST_Windows,"BB"},
			{BuiltInCategory.OST_Railings,"BB"},
			{BuiltInCategory.OST_StairsLandings,"BB"},
			{BuiltInCategory.OST_StairsRuns,"BB"},
			{BuiltInCategory.OST_Stairs,"BB"},
			{BuiltInCategory.OST_CableTrayFitting,"BN"},
			{BuiltInCategory.OST_CableTray,"BN"},
			{BuiltInCategory.OST_Conduit,"BN"},
			{BuiltInCategory.OST_ConduitFitting,"BN"},
			{BuiltInCategory.OST_ElectricalFixtures,"BN"},
			{BuiltInCategory.OST_ElectricalEquipment,"BN"},
			{BuiltInCategory.OST_LightingDevices,"BN"},
			{BuiltInCategory.OST_LightingFixtures,"BN"}
		};
//		public void Hierarchy()
//		{
//			Document doc = this.ActiveUIDocument.Document;		
//			string message=string.Empty;			
//			List<ParameterProperty> parameters=new List<ParameterProperty>{
//				new ParameterProperty("Discipline","BB,BN,BU,BJ,BK",Behavior.Common),
//				new ParameterProperty("Discipline","BH",Behavior.Combine,"Category"),
//				new ParameterProperty("Discipline Code","BB,BN,BU,BJ,BK",Behavior.Common),
//				new ParameterProperty("Discipline Code","BH",Behavior.Combine,"Category"),
//				new ParameterProperty("AGC_DESC","BH,BN,BU,BJ,BK",Behavior.Unique),
//				new ParameterProperty("AGC_DESC","BB",Behavior.Common),
//				new ParameterProperty("LV00_WBS Area Code","BB",Behavior.Common),
//				new ParameterProperty("LV00_WBS Unit Code","BB",Behavior.Common),
//				new ParameterProperty("LV01_Object Code","BB",Behavior.Common),
//				new ParameterProperty("LV01_Sequence Number","BB",Behavior.Common),
//				new ParameterProperty("LV02_Object Code","BB",Behavior.Combine),
//				new ParameterProperty("LV02_Sequence Number","BB",Behavior.Sequence,"Object Code"),
//				new ParameterProperty("PTY_04","BB",Behavior.Common),
//				new ParameterProperty("CWA","BB",Behavior.Common),
//				new ParameterProperty("Main Item Tag","BB",Behavior.CombineCommon,"LV00_WBS Area Code-LV01_Object Code-LV01_Sequence Number"),
//				new ParameterProperty("AGC_TYPE","BH,BB,BN,BU,BJ,BK",Behavior.Combine,"Category,Zone Code,Object Code"),
//				new ParameterProperty("SITE","BB",Behavior.CombineCommon,"001-LV00_WBS Unit Code-CWA-Discipline Code"),
//				new ParameterProperty("SITE","BN,BU,BJ,BK",Behavior.CombineCommon,"001-Process Unit Code-Construction Area-Discipline Code"),
//				new ParameterProperty("SITE","BH",Behavior.Combine,"001-Process Unit Code-Construction Area-Discipline Code"),
//				new ParameterProperty("ZONE","BB",Behavior.CombineCommon,"CWA-Discipline Code/Main Item Tag"),
//				new ParameterProperty("ZONE","BN,BU,BJ,BK",Behavior.CombineCommon,"Construction Area-Discipline Code/Zone Code"),
//				new ParameterProperty("ZONE","BH",Behavior.Combine,"Construction Area-Discipline Code/Zone Code"),
//				new ParameterProperty("STRU","BB",Behavior.CombineCommon,"ZONE/PTY_04"),
//				new ParameterProperty("STRU","BU,BH",Behavior.Combine,"ZONE/"),
//				new ParameterProperty("AGC_DOC","BB",Behavior.CombineCommon,"Main Item Tag-PTY_04"),
//				new ParameterProperty("SBFR","BB",Behavior.Combine,"STRU/LV02_Object Code"),
//				new ParameterProperty("FRMW","BB",Behavior.Combine,"SBFR/LV02_Sequence Number"),
//				new ParameterProperty("SBFR","BH,BU",Behavior.Combine,"STRU/LV02_Object Code"),
//				new ParameterProperty("FRMW","BH,BU",Behavior.Combine,"SBFR/LV02_Sequence Number")
//			};
//			string discipline=doc.Title.Substring(doc.Title.IndexOf("-")+1,2);
//			ElementFilter mcf=new ElementMulticategoryFilter(Categories.Where(x=>x.Value.Contains(discipline)).Select(x=>x.Key).ToList());
//			FilteredElementCollector elems = new FilteredElementCollector(doc).WhereElementIsNotElementType().WherePasses(mcf);
//			Dictionary<string,ParameterProperty> guids=SharedParameters(parameters);
//			Dictionary<string,ParameterProperty> common=guids.Where(x=>x.Value.Behavior==Behavior.Common).ToDictionary(x=>x.Key,x=>x.Value);
//			int i=0;
//			foreach (Element el in elems) {
//			foreach (Parameter par in el.Parameters) {
//					string n=par.Definition.Name;
//					if(!par.IsReadOnly&&par.IsShared&&par.HasValue&&par.StorageType==StorageType.String&&par.AsString()!=""&&guids.ContainsKey(n)){
//						guids[n].Selector.Add("",par.AsString());
//						i++;
//					}
//					if(common.Count==i)break;
//			}	
//			}
//			Dictionary<string,ParameterProperty> combinecommon=guids.Where(x=>x.Value.Behavior==Behavior.CombineCommon).ToDictionary(x=>x.Key,x=>x.Value);
//			Dictionary<string,ParameterProperty> cc=combinecommon;
//			foreach (KeyValuePair<string,ParameterProperty> kvp in common) {
//				foreach (KeyValuePair<string,ParameterProperty> kvpcc in cc){
//					kvp.Value.Argument.Replace(kvpcc.Key,kvpcc.Value.Value());
//					if (!combinecommon.ContainsKey(kvpcc.Key)) {
//						combinecommon.Add(kvp.Key,kvp.Value);
//					}
//				}
//			}
//			Transaction trans=new Transaction(doc);
//			trans.Start("Hierarchy "+discipline);
//			foreach (Element element in elems) {
//				Dictionary<string,Parameter> elparams=new Dictionary<string,Parameter>();
//				foreach (KeyValuePair<string,ParameterProperty> kvp in guids) {
//					string pn=kvp.Key;
//					Parameter par =element.get_Parameter(guids[pn].GUID);
//					if(par==null)message+="Missing Reference: "+pn+" in "+element.Name+"\n";
//					else{ elparams.Add(pn,par);
//						if (!par.HasValue)par.Set(string.Empty);
//					}
//				}
//				elparams["SITE"].Set("001-"+elparams["LV00_WBS Unit Code"].AsString()+"-"+elparams["CWA"].AsString()+"-"+elparams["Discipline Code"].AsString());
//			}
//			trans.Commit();
//			if(message!=String.Empty)TaskDialog.Show("Messages",message);
//		}
		class GuidChecker
		{
			public Guid Original{get;set;}
			public List<Guid> Collisions{get;set;}
			public bool Exist{get;set;}
			public GuidChecker(Guid original, Guid target)
			{
				Collisions=new List<Guid>();
				Original=original;
				Add(target);
			}
			public void Add(Guid target){
				if(Original==target)Exist=true;
				else Collisions.Add(target);
			}
		}
		class GuidFixer
		{
			public string Name{get;set;}
			public ElementBinding Binding{get;set;}
			public Guid Original{get;set;}
			public Dictionary<Guid,ElementBinding> Collisions{get;set;}
			public bool Exist{get;set;}
			public GuidFixer(string name,Guid original, ElementBinding binding)
			{
				Collisions=new Dictionary<Guid,ElementBinding>();
				Name=name;
				Original=original;
				Binding=binding;
			}
			public void Add(Guid target, ElementBinding binding){
				if(Original==target)Exist=true;
				else{
					foreach (KeyValuePair<Guid,ElementBinding> kvp in new Dictionary<Guid,ElementBinding>(Collisions)) {
						if (binding is InstanceBinding||binding.GetType().Equals(kvp.Value.GetType()))Collisions.Add(target,binding);
						else if (!binding.GetType().Equals(kvp.Value.GetType()))Collisions.Remove(kvp.Key);
					}
				}
			}
		}
//		public void SharedParameterFixer()
//		{
//		 	string message=string.Empty;
//			Document doc=this.ActiveUIDocument.Document;
//			BindingMap map = doc.ParameterBindings;
//			DefinitionBindingMapIterator it= map.ForwardIterator();
//			DefinitionFile spFile = this.Application.OpenSharedParameterFile();
//			Dictionary<Guid,GuidFixer> results=new Dictionary<Guid, GuidFixer>();
//			Dictionary<int,ElementBinding> defs=new Dictionary<int,ElementBinding>();
//			it.Reset();
//			if(spFile!=null)
//				while(it.MoveNext())						
//				foreach (DefinitionGroup dG in spFile.Groups)
//						foreach (ExternalDefinition eD in dG.Definitions)
//							if (eD.Name==it.Key.Name){
//				Guid guid=eD.GUID;
//				if (!results.ContainsKey(guid))results.Add(guid,new GuidFixer(eD.Name,guid,(ElementBinding)it.Current));
//				defs.Add(((InternalDefinition)it.Key).Id.IntegerValue,(ElementBinding)it.Current);
//				}
//			List<SharedParameterElement> spelist=new List<SharedParameterElement>(new FilteredElementCollector(doc).OfClass(typeof(SharedParameterElement)).Cast<SharedParameterElement>());
//			foreach (SharedParameterElement spe in spelist) {
//				Guid guid=spe.GuidValue;
//				int id=spe.GetDefinition().Id.IntegerValue;
//				if (defs.ContainsKey(id)) {
//				foreach (KeyValuePair<Guid,GuidFixer> kvp in results)
//					if (kvp.Value.Name==spe.Name)kvp.Value.Add(guid,defs[id]);
//				}
//				defs.Remove(id);
//				}
//			foreach (KeyValuePair<Guid,GuidFixer> kvp in new Dictionary<Guid,GuidFixer>(results))
//				if (kvp.Value.Exist&&kvp.Value.Collisions.Count==0)results.Remove(kvp.Key);
//			foreach (KeyValuePair<Guid,GuidFixer> kvp in results) {
//				foreach (Guid g in kvp.Value.Collisions) {
//					message+=g.ToString()+"\n";
//				}
//			}
//			TaskDialog.Show("msg", message);
//}
}
}
