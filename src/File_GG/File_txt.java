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
	    * 读取 .txt 后缀文件，将所有内容存放到 Arraylist变量中
	    * @param File_Path
	    *        文件的绝对路径
	    * @param Arraylist
	    * 		  全局变量,定义示例:<br>
	    *        static ArrayList <String> LIST=new ArrayList<String>();
	*/
	public static void read(String File_Path,ArrayList <String> Arraylist)
	{
		String neirong=" ";
		try
		{
			BufferedReader reader=new BufferedReader(new FileReader(File_Path)); 
			while (neirong!= null) //null时表示读到文件末
			{
				neirong=reader.readLine();  //读一行数据
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
	    * 读取 .txt 后缀文件，将所有内容存放到 Arraylist变量中,若某行的内容是NO_read变量的字符串跳过不读取
	    * (多用于空字符串——NO_read="")
	    * @param File_Path
	    *        文件的绝对路径
	    * @param Arraylist
	    * 		  全局变量,定义示例:<br>
	    *        static ArrayList <String> LIST=new ArrayList<String>();
	    * @param NO_read
	    * 		  遇到此字符串跳过不读取
	*/
	public static void read(String File_Path,ArrayList <String> Arraylist,String NO_read)
	{
		String neirong=" ";
		try
		{
			BufferedReader reader=new BufferedReader(new FileReader(File_Path)); 
			while (neirong!= null) //null时表示读到文件末
			{
				neirong=reader.readLine();  //读一行数据
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
	    * 向 .txt 后缀文件写入 Arraylist变量的所有内容,若写入前该文件存在，则删除后再创建
	    * @param File_Path
	    *        文件的绝对路径
	    * @param Arraylist
	    * 		  全局变量,定义示例:<br>
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
	    * 向 已有的.txt 后缀文件写入 string变量的内容,在文件的最后一行写入数据,若文件不存在则先创建再写入
	    * @param File_Path
	    *        文件的绝对路径
	    * @param string
	    * 		 String变量,存放着要写入的数据
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
