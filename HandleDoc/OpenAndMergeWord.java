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

        uniteDoc(list,"D:/docname3.doc"); 	//�ϲ�doc�ĵ�
        
        opendoc("D:/docname3.doc");			//��doc�ĵ�
	}
	public static void uniteDoc(List fileList, String savepaths) {  
        if (fileList.size() == 0 || fileList == null) {  
            return;  
        }  
        //��word  
        ActiveXComponent app = new ActiveXComponent("KWPS.Application");
//		      �����°�Ķ�,��wps ��"wps.application" ��Ϊ "KWPS.Application"  ��word�򲻱䣬��ȻΪ "word.application"
//        �ҽ���ŷ��֣����°棬CLASS���ֶ��ı��ˡ���ԭ����CLASSǰ�����K��	(2014��4��)
//        KWPS.Application
//        KET.Application
//        KWPP.Application
//        ������Щ���������3.0�ˡ�
        try {  
            // ����word���ɼ�  
            app.setProperty("Visible", new Variant(false));  
            //���documents����  
            Object docs = app.getProperty("Documents").toDispatch();  
            //�򿪵�һ���ļ�  
            Object doc = Dispatch  
                .invoke(  
                        (Dispatch) docs,  
                        "Open",  
                        Dispatch.Method,  
                        new Object[] { (String) fileList.get(0),  
                                new Variant(false), new Variant(true) },  
                        new int[3]).toDispatch();  
            //׷���ļ�  
            for (int i = 1; i < fileList.size(); i++) {  
                Dispatch.invoke(app.getProperty("Selection").toDispatch(),  
                    "insertFile", Dispatch.Method, new Object[] {  
                            (String) fileList.get(i), "",  
                            new Variant(false), new Variant(false),  
                            new Variant(false) }, new int[3]);  
            }  
            //�����µ�word�ļ�  
            Dispatch.invoke((Dispatch) doc, "SaveAs", Dispatch.Method,  
                new Object[] { savepaths, new Variant(1) }, new int[3]);  
            Variant f = new Variant(false);  
            Dispatch.call((Dispatch) doc, "Close", f);  
        } catch (Exception e) {  
            throw new RuntimeException("�ϲ�word�ļ�����.ԭ��:" + e);  
        } finally {  
            app.invoke("Quit", new Variant[] {});  
        }  
    }  
	
	public static void opendoc(String savepaths){
		 ActiveXComponent app = new ActiveXComponent("KWPS.Application");//����word  
	        try {  
	            // ����word���ɼ�  
	            app.setProperty("Visible", new Variant(true));  
	            //���documents����  
	            Object docs = app.getProperty("Documents").toDispatch();  
	            //�򿪵�һ���ļ�  
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
