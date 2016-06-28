package eclipseex.views;


import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Composite;
import org.eclipse.ui.part.*;
import org.eclipse.jface.viewers.*;
import org.eclipse.swt.graphics.Color;
import org.eclipse.swt.graphics.Image;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.ole.win32.OLE;
import org.eclipse.swt.ole.win32.OleAutomation;
import org.eclipse.swt.ole.win32.OleClientSite;
import org.eclipse.swt.ole.win32.OleFrame;
import org.eclipse.swt.ole.win32.Variant;
import org.eclipse.jface.action.*;
import org.eclipse.jface.dialogs.MessageDialog;
import org.eclipse.jface.util.IPropertyChangeListener;
import org.eclipse.jface.util.PropertyChangeEvent;
import org.eclipse.ui.*;
import org.eclipse.swt.widgets.Menu;
import org.eclipse.swt.widgets.Text;
import org.eclipse.swt.SWT;
import org.eclipse.swt.SWTError;
import org.eclipse.swt.events.SelectionAdapter;
import org.eclipse.swt.events.SelectionEvent;


/**
 * This sample class demonstrates how to plug-in a new
 * workbench view. The view shows data obtained from the
 * model. The sample creates a dummy model on the fly,
 * but a real implementation would connect to the model
 * available either in this or another plug-in (e.g. the workspace).
 * The view is connected to the model using a content provider.
 * <p>
 * The view uses a label provider to define how model
 * objects should be presented in the view. Each
 * view can present the same model objects using
 * different labels and icons, if needed. Alternatively,
 * a single label provider can be shared between views
 * in order to ensure that objects of the same type are
 * presented in the same way everywhere.
 * <p>
 */

public class EclipseExView extends ViewPart {

	/**
	 * The ID of the view as specified by the extension.
	 */
	public static final String ID = "eclipseex.views.EclipseExView";

	private TableViewer viewer;
	private Button Tangram1;
	private Button Tangram2;
	private Text text;
	 

	class ViewLabelProvider extends LabelProvider implements ITableLabelProvider {
		public String getColumnText(Object obj, int index) {
			return getText(obj);
		}
		public Image getColumnImage(Object obj, int index) {
			return getImage(obj);
		}
		public Image getImage(Object obj) {
			return PlatformUI.getWorkbench().getSharedImages().getImage(ISharedImages.IMG_OBJ_ELEMENT);
		}
	}

/**
	 * The constructor.
	 */

	public EclipseExView() 
	{
		addPartPropertyListener(new IPropertyChangeListener() 
		{
			public void propertyChange(PropertyChangeEvent event) 
			{
			}
		});
	}


/**
 * This is a callback that will allow us
 * to create the viewer and initialize it.
 */
private String strXml = "<tangram framme=\"\">" +
        "<window>" +
            "<node name=\"Start_S0001\" id=\"TangramTab\" cnnid=\"TangramTabbedWnd.TabbedComponent.1\" activepage=\"3\" style=\"18\">" +
                "<node name=\"Page1\" caption=\"TangramEclipseView\" id=\"hostview\" cnnid=\"\" />" +
                "<node name=\"Page2\" caption=\"TangramPage2\" id=\"\" cnnid=\"TangramCtrl.UserCtrl,TangramCtrl\" />" +
                "<node name=\"Page3\" caption=\"TangramPage3\" id=\"\" cnnid=\"\" />" +
            "</node>" +
        "</window>" +
    "</tangram>";
private String strXml2 = "<tangram framme=\"\">" +
        "<window>" +
            "<node name=\"Start_S0001\" id=\"TangramTab\" cnnid=\"TangramTabbedWnd.TabbedComponent.1\" activepage=\"3\" style=\"39\">" +
                "<node name=\"Page1\" caption=\"TangramInnerView\" id=\"hostview\" cnnid=\"\" />" +
                "<node name=\"Page2\" caption=\"App Page2\" id=\"\" cnnid=\"TangramCtrl.UserCtrl,TangramCtrl\" />" +
                "<node name=\"Page3\" caption=\"App Page3\" id=\"\" cnnid=\"\" />" +
            "</node>" +
        "</window>" +
    "</tangram>";
private OleClientSite site;
static final int Extend = 103;
static final int TangramHandle = 2;
private OleAutomation auto = null;
public void createPartControl(Composite parent) {
	Variant varTextView = new Variant("TextView");
	Variant varView = new Variant("EclipseView");
	Variant varKey = new Variant("test");
	//xObj = new Tangram.ActiveX.TangramComObj("tangram.tangram.1");
	try {
	      OleFrame frame = new OleFrame(parent, SWT.NONE);
	      frame.setVisible(false);
	      site = new OleClientSite(frame, SWT.NONE, "Tangram.Ctrl.1");
	      //site.setBounds(0, 0, 1, 1);
	      site.doVerb(OLE.OLEIVERB_INPLACEACTIVATE);
	      auto = new OleAutomation(site);
	      frame.setSize(0, 0);
	      site.setBounds(0, 0, 1, 1);
	      } catch (SWTError e) {
	      System.out.println("Unable to open activeX control");
	      return;
	    }
	GridLayout gridLayout = new GridLayout();
	gridLayout.numColumns=1;
	parent.setLayout(gridLayout);
	Composite Panel=new Composite(parent,SWT.NONE);
	GridLayout gl_Panel = new GridLayout();
	gl_Panel.numColumns=2;
	Panel.setLayout(gl_Panel);
	Panel.setSize(30, 30);
	Tangram1 = new Button(Panel, SWT.NONE);
	Tangram1.setText("Tangram1");

	Tangram1.addSelectionListener(new SelectionAdapter(){
		public void widgetSelected(SelectionEvent e){
			String tContent = text.getText();
		    auto.invoke(Extend, new Variant[] { new Variant("TopView"), new Variant("button1"),new Variant(tContent) });
			}
	});


	Tangram2 = new Button(Panel, SWT.NONE);
	Tangram2.setText("Tangram2");
	//button2.setLayoutData();
	Tangram2.addSelectionListener(new SelectionAdapter(){
		public void widgetSelected(SelectionEvent e){
			String tContent = text.getText();
		    auto.invoke(Extend, new Variant[] { new Variant("TopView"), new Variant("button2"),new Variant(tContent) });
			}
	});
	
	GridData g2 = new GridData();
	
	g2.heightHint=400;
	g2.widthHint=600;
	g2.grabExcessHorizontalSpace=true;
	g2.grabExcessVerticalSpace=true;
	g2.horizontalAlignment=GridData.FILL;
	
	Composite subComposite=new Composite(parent,SWT.NONE);
	subComposite.setLayoutData(new GridData(SWT.LEFT, SWT.FILL, false, false, 1, 1));
	GridLayout gridLayout1 = new GridLayout();
	gridLayout1.numColumns=1;
	subComposite.setLayout(gridLayout1);
	
	text = new org.eclipse.swt.widgets.Text(subComposite,SWT.MULTI|SWT.V_SCROLL|SWT.H_SCROLL);
	text.setLayoutData(g2);
	text.setText("hello word");

	text.setBackground( new Color(null, 205,205,205));

    auto.setProperty(TangramHandle, new Variant[] { varView, new Variant((long)parent.handle)});
    auto.setProperty(TangramHandle, new Variant[] { varTextView, new Variant((long)text.handle)});
    auto.invoke(Extend, new Variant[] { varView, varKey,new Variant(strXml)});
    auto.invoke(Extend, new Variant[] { varTextView, varKey,new Variant(strXml2) });
}

/**
 * Passing the focus request to the viewer's control.
 */
public void setFocus() {
	viewer.getControl().setFocus();
	}
}
