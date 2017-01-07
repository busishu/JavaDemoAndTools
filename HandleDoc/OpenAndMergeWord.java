package Word;

import java.util.ArrayList;
import java.util.List;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class OpenAndMergeWord {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		List list  = new ArrayList();  
        String file1= "D:/docname1"+".doc";  
        String file2= "D:/docname2"+".doc";  
 
        list.add(file2);  
        list.add(file1);  

        uniteDoc(list,"D:/docname3.doc"); 	//合并doc文档
        
        opendoc("D:/docname3.doc");			//打开doc文档
	}
	public static void uniteDoc(List fileList, String savepaths) {  
        if (fileList.size() == 0 || fileList == null) {  
            return;  
        }  
        //打开word  
        ActiveXComponent app = new ActiveXComponent("KWPS.Application");
//		      由于新版改动,打开wps 从"wps.application" 改为 "KWPS.Application"  打开word则不变，仍然为 "word.application"
//        我今天才发现，四月版，CLASS名字都改变了。在原来的CLASS前面加了K：	(2014年4月)
//        KWPS.Application
//        KET.Application
//        KWPP.Application
//        现在这些组件升级到3.0了。
        try {  
            // 设置word不可见  
            app.setProperty("Visible", new Variant(false));  
            //获得documents对象  
            Object docs = app.getProperty("Documents").toDispatch();  
            //打开第一个文件  
            Object doc = Dispatch  
                .invoke(  
                        (Dispatch) docs,  
                        "Open",  
                        Dispatch.Method,  
                        new Object[] { (String) fileList.get(0),  
                                new Variant(false), new Variant(true) },  
                        new int[3]).toDispatch();  
            //追加文件  
            for (int i = 1; i < fileList.size(); i++) {  
                Dispatch.invoke(app.getProperty("Selection").toDispatch(),  
                    "insertFile", Dispatch.Method, new Object[] {  
                            (String) fileList.get(i), "",  
                            new Variant(false), new Variant(false),  
                            new Variant(false) }, new int[3]);  
            }  
            //保存新的word文件  
            Dispatch.invoke((Dispatch) doc, "SaveAs", Dispatch.Method,  
                new Object[] { savepaths, new Variant(1) }, new int[3]);  
            Variant f = new Variant(false);  
            Dispatch.call((Dispatch) doc, "Close", f);  
        } catch (Exception e) {  
            throw new RuntimeException("合并word文件出错.原因:" + e);  
        } finally {  
            app.invoke("Quit", new Variant[] {});  
        }  
    }  
	
	public static void opendoc(String savepaths){
		 ActiveXComponent app = new ActiveXComponent("KWPS.Application");//启动word  
	        try {  
	            // 设置word不可见  
	            app.setProperty("Visible", new Variant(true));  
	            //获得documents对象  
	            Object docs = app.getProperty("Documents").toDispatch();  
	            //打开第一个文件  
	            Object doc = Dispatch  
	                .invoke(  
	                        (Dispatch) docs,  
	                        "Open",  
	                        Dispatch.Method,  
	                        new Object[] { savepaths,  
	                                new Variant(false), new Variant(true) },  
	                        new int[3]).toDispatch();  
	            }catch(Exception e){
	            	e.printStackTrace();
	            }
	}

}
