package File_GG;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;

public class File_txt {

	/**	
	    * ��ȡ .txt ��׺�ļ������������ݴ�ŵ� Arraylist������
	    * @param File_Path
	    *        �ļ��ľ���·��
	    * @param Arraylist
	    * 		  ȫ�ֱ���,����ʾ��:<br>
	    *        static ArrayList <String> LIST=new ArrayList<String>();
	*/
	public static void read(String File_Path,ArrayList <String> Arraylist)
	{
		String neirong=" ";
		try
		{
			BufferedReader reader=new BufferedReader(new FileReader(File_Path)); 
			while (neirong!= null) //nullʱ��ʾ�����ļ�ĩ
			{
				neirong=reader.readLine();  //��һ������
				if(neirong == null)
				{
					continue;
				}
				Arraylist.add(neirong);
			}
			
			reader.close();
		}
		catch(IOException e)
		{
			e.printStackTrace();
		}
	}
	
	/**	
	    * ��ȡ .txt ��׺�ļ������������ݴ�ŵ� Arraylist������,��ĳ�е�������NO_read�������ַ�����������ȡ
	    * (�����ڿ��ַ�������NO_read="")
	    * @param File_Path
	    *        �ļ��ľ���·��
	    * @param Arraylist
	    * 		  ȫ�ֱ���,����ʾ��:<br>
	    *        static ArrayList <String> LIST=new ArrayList<String>();
	    * @param NO_read
	    * 		  �������ַ�����������ȡ
	*/
	public static void read(String File_Path,ArrayList <String> Arraylist,String NO_read)
	{
		String neirong=" ";
		try
		{
			BufferedReader reader=new BufferedReader(new FileReader(File_Path)); 
			while (neirong!= null) //nullʱ��ʾ�����ļ�ĩ
			{
				neirong=reader.readLine();  //��һ������
				if(neirong == null || neirong.equals(NO_read))
				{
					continue;
				}
				Arraylist.add(neirong);
			}
			
			reader.close();
		}
		catch(IOException e)
		{
			e.printStackTrace();
		}
	}
	
	/**	
	    * �� .txt ��׺�ļ�д�� Arraylist��������������,��д��ǰ���ļ����ڣ���ɾ�����ٴ���
	    * @param File_Path
	    *        �ļ��ľ���·��
	    * @param Arraylist
	    * 		  ȫ�ֱ���,����ʾ��:<br>
	    *        static ArrayList <String> LIST=new ArrayList<String>();
	*/
	public static void write(String File_Path,ArrayList <String> Arraylist)
	{
		FileWriter fw = null;
		BufferedWriter bw = null;
		try
		{
			File txt=new File(File_Path);
			if(txt.exists())
			{
				txt.delete();
			}
			txt.createNewFile();
			
			fw = new FileWriter(File_Path, true);
			bw = new BufferedWriter(fw);
			
			for(int i=0;i<Arraylist.size();i++)
			{
				bw.write(Arraylist.get(i)+"\r\n");
				bw.flush();
			}

			bw.close();
			fw.close();
		}
		catch(IOException e)
		{
			e.printStackTrace();
		}
	}
	
	/**	
	    * �� ���е�.txt ��׺�ļ�д�� string����������,���ļ������һ��д������,���ļ����������ȴ�����д��
	    * @param File_Path
	    *        �ļ��ľ���·��
	    * @param string
	    * 		 String����,�����Ҫд�������
	*/
	public static void write(String File_Path,String string)
	{
		FileWriter fw = null;
		BufferedWriter bw = null;
		try
		{
			/*
			File txt=new File(File_Path);
			if(txt.exists())
			{
				txt.delete();
			}
			txt.createNewFile();
			*/
			File txt=new File(File_Path);
			if(!txt.exists())
			{
				txt.createNewFile();
			}
			
			fw = new FileWriter(File_Path, true);
			bw = new BufferedWriter(fw);
			bw.write(string+"\r\n");
			bw.flush();

			bw.close();
			fw.close();
		}
		catch(IOException e)
		{
			e.printStackTrace();
		}
	}
	
	
}
