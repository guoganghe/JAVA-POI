package File_GG;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.ss.util.CellRangeAddress;


public class File_excel {
	
	/**
	    * 通过工作表的序号打开工作表,能读取合并单元格的值
	    * @param File_Path
	    *        文件的绝对路径
	    * @param sheet_num
	    * 		 工作簿中的工作表序号，第一个工作表记为0
	    * @param hang
	    * 		 工作表中的行数,EXCEL文件中的第1行记为0
	    * @param lie
	    * 		 工作表中的列数,EXCEL文件中的第A行记为0
	    * @return String
	*/
	public static String read(String File_Path,int sheet_num,int hang,int lie)
	{
		String str="";
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls后缀           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel的sheet
				HSSFRow row;             //excel的行
				HSSFCell cell;           //excel的列
				
				sheet = wb.getSheetAt(sheet_num);
				int maxrow=sheet.getLastRowNum();  //获得该sheet工作表上最大物理行
				if(hang>maxrow)
				{
					//System.out.println("超过最大有效物理行数:"+hang+",最大行数为:"+maxrow);
					wb.close();
					return null;
				}

				row = sheet.getRow(hang);
				if(row==null)
				{
					wb.close();
					return null;
				}
				/*
				int maxcell=row.getLastCellNum();  //获得该sheet工作表上最大列
				if(lie>maxcell)
				{
					//System.out.println(hang+"行--"+"超过最大有效物理列数:"+lie+",最大列数为:"+maxcell);
					wb.close();
					return null;
				}*/
				cell=row.getCell(lie);

				if(cell==null || cell.toString()==(""))
				{
					//System.out.println("读到空");
					
					ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
					//判断是否是合并单元格
					int sheetmergerCount = sheet.getNumMergedRegions();
					if(sheetmergerCount==0)
					{
						wb.close();
						return null;
					}
					
					for (int i = 0; i < sheetmergerCount; i++)
					{
						  // 获得合并单元格加入list中
						  CellRangeAddress ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //返回指定索引合并后的区域
						  LIST.add(ca);
					}	
					
					int firstC = 0;
					int lastC = 0;
					int firstR = 0;
					int lastR = 0;
					
					int run_i;
					for (run_i=0;run_i<LIST.size();run_i++)
					{
						CellRangeAddress ca=LIST.get(run_i);
						// 获得合并单元格的起始行, 结束行, 起始列, 结束列
						firstC = ca.getFirstColumn();
						lastC = ca.getLastColumn();
						firstR = ca.getFirstRow();
						lastR = ca.getLastRow();
						if (lie <= lastC&& lie>= firstC)
						{
							if (hang <= lastR && hang >= firstR)
							{
								//合并单元格的值，读取合并区域的首行首列
								row = sheet.getRow(firstR);
								cell=row.getCell(firstC);
								break;
							}
						}
					}
					
					//不是合并单元格,其单元格内没有内容
					if(run_i>=LIST.size())
					{
						wb.close();
						return null;
					}
					
					//合并区域的首行首列无内容
					if(cell==null || cell.toString()==(""))
					{
						wb.close();
						return null;
					}
				}
				switch (cell.getCellType()) 
				{
					case HSSFCell.CELL_TYPE_STRING: //字符串类型
						str=cell.getStringCellValue();
						break;
					case HSSFCell.CELL_TYPE_NUMERIC: //数值类型
						/*
					case Cell.CELL_TYPE_NUMERIC:  
						cell.setCellType(Cell.CELL_TYPE_STRING);  
						literal = cell.getStringCellValue();  
						// POI Bug 
						literal = Double.toString(cell.getNumericCellValue());  
						literal = new DataFormatter().formatCellValue(cell);  
						break; 
                         */
						
						
						double qudian;
						qudian=cell.getNumericCellValue();
						//System.out.println(qudian);
						str=String.format("%.0f", qudian);
						//codehex=String.valueOf(hex.getNumericCellValue()); 
						break;
				}
				wb.close();
				
			}
			else         //.xlsx后缀        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel的sheet
				XSSFRow row;             //excel的行
				XSSFCell cell;           //excel的列
				
				sheet = xwb.getSheetAt(sheet_num);
				
				int maxrow=sheet.getLastRowNum();  //获得该sheet工作表上最大物理行
				if(hang>maxrow)
				{
					//System.out.println("超过最大有效物理行数:"+hang+",最大行数为:"+maxrow);
					xwb.close();
					return null;
				}

				row = sheet.getRow(hang);
				if(row==null)
				{
					xwb.close();
					return null;
				}
				/*
				int maxcell=row.getLastCellNum();  //获得该sheet工作表上最大物理列
				if(lie>maxcell)
				{
					//System.out.println(hang+"行--"+"超过最大有效物理列数:"+lie+",最大列数为:"+maxcell);
					xwb.close();
					return null;
				}*/
				
				cell=row.getCell(lie);
				if(cell==null || cell.toString()==(""))
				{
					//System.out.println("读到空");
					
					ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
					//判断是否是合并单元格
					int sheetmergerCount = sheet.getNumMergedRegions();
					if(sheetmergerCount==0)
					{
						xwb.close();
						return null;
					}
					
					for (int i = 0; i < sheetmergerCount; i++)
					{
						  // 获得合并单元格加入list中
						  CellRangeAddress ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //返回指定索引合并后的区域
						  LIST.add(ca);
					}	
					
					int firstC = 0;
					int lastC = 0;
					int firstR = 0;
					int lastR = 0;
					
					int run_i;
					for (run_i=0;run_i<LIST.size();run_i++)
					{
						CellRangeAddress ca=LIST.get(run_i);
						// 获得合并单元格的起始行, 结束行, 起始列, 结束列
						firstC = ca.getFirstColumn();
						lastC = ca.getLastColumn();
						firstR = ca.getFirstRow();
						lastR = ca.getLastRow();
						if (lie <= lastC&& lie>= firstC)
						{
							if (hang <= lastR && hang >= firstR)
							{
								//合并单元格的值，读取合并区域的首行首列
								row = sheet.getRow(firstR);
								cell=row.getCell(firstC);
								break;
							}
						}
					}
					
					//不是合并单元格,其单元格内没有内容
					if(run_i>=LIST.size())
					{
						xwb.close();
						return null;
					}
					
					//合并区域的首行首列无内容
					if(cell==null || cell.toString()==(""))
					{
						xwb.close();
						return null;
					}
				}
				switch (cell.getCellType()) 
				{
					case XSSFCell.CELL_TYPE_STRING: //字符串类型
						str=cell.getStringCellValue();
						break;
					case XSSFCell.CELL_TYPE_NUMERIC: //数值类型
						double qudian;
						qudian=cell.getNumericCellValue();
						str=String.format("%.0f", qudian);
						//codehex=String.valueOf(hex.getNumericCellValue()); 
						break;
				}
				xwb.close();
				
			}
			
	    }
		catch (IOException e)
		{
			e.printStackTrace();
		}
		return str;

	}
	
	/**	
	    * 通过工作表的名称打开工作表,能读取合并单元格的值
	    * @param File_Path
	    *        文件的绝对路径
	    * @param sheet_name
	    * 		 工作簿中的工作表名称
	    * @param hang
	    * 		 工作表中的行数,EXCEL文件中的第1行记为0
	    * @param lie
	    * 		 工作表中的列数,EXCEL文件中的第A行记为0
	    * @return String
	*/
	public static String read(String File_Path,String sheet_name,int hang,int lie)
	{
		String str="";
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls后缀           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel的sheet
				HSSFRow row;             //excel的行
				HSSFCell cell;           //excel的列
				
				sheet = wb.getSheet(sheet_name);
				
				int maxrow=sheet.getLastRowNum();  //获得该sheet工作表上最大物理行
				if(hang>maxrow)
				{
					//System.out.println("超过最大有效物理行数:"+hang+",最大行数为:"+maxrow);
					wb.close();
					return null;
				}

				row = sheet.getRow(hang);
				if(row==null)
				{
					wb.close();
					return null;
				}
				/*
				int maxcell=row.getLastCellNum();  //获得该sheet工作表上最大物理列
				if(lie>maxcell)
				{
					//System.out.println(hang+"行--"+"超过最大有效物理列数:"+lie+",最大列数为:"+maxcell);
					wb.close();
					return null;
				}*/
				cell=row.getCell(lie);
				
				if(cell==null || cell.toString()==(""))
				{
					//System.out.println("读到空");
					
					ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
					//判断是否是合并单元格
					int sheetmergerCount = sheet.getNumMergedRegions();
					if(sheetmergerCount==0)
					{
						wb.close();
						return null;
					}
					
					for (int i = 0; i < sheetmergerCount; i++)
					{
						  // 获得合并单元格加入list中
						  CellRangeAddress ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //返回指定索引合并后的区域
						  LIST.add(ca);
					}	
					
					int firstC = 0;
					int lastC = 0;
					int firstR = 0;
					int lastR = 0;
					
					int run_i;
					for (run_i=0;run_i<LIST.size();run_i++)
					{
						CellRangeAddress ca=LIST.get(run_i);
						// 获得合并单元格的起始行, 结束行, 起始列, 结束列
						firstC = ca.getFirstColumn();
						lastC = ca.getLastColumn();
						firstR = ca.getFirstRow();
						lastR = ca.getLastRow();
						if (lie <= lastC&& lie>= firstC)
						{
							if (hang <= lastR && hang >= firstR)
							{
								//合并单元格的值，读取合并区域的首行首列
								row = sheet.getRow(firstR);
								cell=row.getCell(firstC);
								break;
							}
						}
					}
					
					//不是合并单元格,其单元格内没有内容
					if(run_i>=LIST.size())
					{
						wb.close();
						return null;
					}
					
					//合并区域的首行首列无内容
					if(cell==null || cell.toString()==(""))
					{
						wb.close();
						return null;
					}
				}
				switch (cell.getCellType()) 
				{
					case HSSFCell.CELL_TYPE_STRING: //字符串类型
						str=cell.getStringCellValue();
						break;
					case HSSFCell.CELL_TYPE_NUMERIC: //数值类型
						double qudian;
						qudian=cell.getNumericCellValue();
						str=String.format("%.0f", qudian);
						//codehex=String.valueOf(hex.getNumericCellValue()); 
						break;
				}
				wb.close();
				
			}
			else         //.xlsx后缀        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel的sheet
				XSSFRow row;             //excel的行
				XSSFCell cell;           //excel的列
				
				sheet = xwb.getSheet(sheet_name);
				
				int maxrow=sheet.getLastRowNum();  //获得该sheet工作表上最大物理行
				if(hang>maxrow)
				{
					//System.out.println("超过最大有效物理行数:"+hang+",最大行数为:"+maxrow);
					xwb.close();
					return null;
				}

				row = sheet.getRow(hang);
				if(row==null)
				{
					xwb.close();
					return null;
				}
				/*
				int maxcell=row.getLastCellNum();  //获得该sheet工作表上最大物理列
				if(lie>maxcell)
				{
					//System.out.println(hang+"行--"+"超过最大有效物理列数:"+lie+",最大列数为:"+maxcell);
					xwb.close();
					return null;
				}*/
				
				cell=row.getCell(lie);
				if(cell==null || cell.toString()==(""))
				{
					//System.out.println("读到空");
					
					ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
					//判断是否是合并单元格
					int sheetmergerCount = sheet.getNumMergedRegions();
					if(sheetmergerCount==0)
					{
						xwb.close();
						return null;
					}
					
					for (int i = 0; i < sheetmergerCount; i++)
					{
						  // 获得合并单元格加入list中
						  CellRangeAddress ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //返回指定索引合并后的区域
						  LIST.add(ca);
					}	
					
					int firstC = 0;
					int lastC = 0;
					int firstR = 0;
					int lastR = 0;
					
					int run_i;
					for (run_i=0;run_i<LIST.size();run_i++)
					{
						CellRangeAddress ca=LIST.get(run_i);
						// 获得合并单元格的起始行, 结束行, 起始列, 结束列
						firstC = ca.getFirstColumn();
						lastC = ca.getLastColumn();
						firstR = ca.getFirstRow();
						lastR = ca.getLastRow();
						if (lie <= lastC&& lie>= firstC)
						{
							if (hang <= lastR && hang >= firstR)
							{
								//合并单元格的值，读取合并区域的首行首列
								row = sheet.getRow(firstR);
								cell=row.getCell(firstC);
								break;
							}
						}
					}
					
					//不是合并单元格,其单元格内没有内容
					if(run_i>=LIST.size())
					{
						xwb.close();
						return null;
					}
					
					//合并区域的首行首列无内容
					if(cell==null || cell.toString()==(""))
					{
						xwb.close();
						return null;
					}
				}
				switch (cell.getCellType()) 
				{
					case XSSFCell.CELL_TYPE_STRING: //字符串类型
						str=cell.getStringCellValue();
						break;
					case XSSFCell.CELL_TYPE_NUMERIC: //数值类型
						double qudian;
						qudian=cell.getNumericCellValue();
						str=String.format("%.0f", qudian);
						//codehex=String.valueOf(hex.getNumericCellValue()); 
						break;
				}
				xwb.close();
				
			}
			
	    }
		catch (IOException e)
		{
			e.printStackTrace();
		}
		return str;

	}
	
	/**	
	    * 获得一个EXCEL文件指定序号的工作表名称<br>
	    * 第一个工作表序号是0
	    * @param File_Path
	    *        文件的绝对路径
	    * @param sheet_num
	    * 		 工作簿中的工作表序号,第一个工作表序号记为0
	    * @return String
	*/
	public static String getsheetname(String File_Path,int sheet_num)
	{
		String sheetname="";
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls后缀           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				sheetname=wb.getSheetName(sheet_num);  //获得sheet名称
				wb.close();
			}
			else                        //.xlsx后缀        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				sheetname=xwb.getSheetName(sheet_num);  //获得sheet名称
				xwb.close();
			}
	    }
		catch (IOException e)
		{
			e.printStackTrace();
		}
		return sheetname;
	}
	
	/**	
	    * 获得一个EXCEL工作簿的工作表数量<br>
	    * 如返回值=2,则有2个工作表——序号为0、1
	    * @param File_Path
	    *        文件的绝对路径
	    * @return int
	*/
	public static int getsheetnum(String File_Path)
	{
		int sheet_num=0;
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls后缀           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				sheet_num=wb.getNumberOfSheets();  //获得sheet数量
				wb.close();
			}
			else                        //.xlsx后缀        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				sheet_num=xwb.getNumberOfSheets();  //获得sheet数量
				xwb.close();
			}
	    }
		catch (IOException e)
		{
			e.printStackTrace();
		}
		return sheet_num;
	}
	
	/**	
	    * 获得一个EXCEL工作簿指定工作表的最大行数,通过工作表序号打开<br>
	    * 如返回值=4(0、1、2、3),则第4行及后面的行上的单元格都没有内容
	    * @param File_Path
	    *        文件的绝对路径
	    * @param sheet_num
	    * 		  工作簿中的工作表序号,第一个工作表序号记为0
	    * @return int
	*/
	public static int getmaxrow(String File_Path,int sheet_num)
	{
		int maxrow=0;
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls后缀           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel的sheet
				sheet = wb.getSheetAt(sheet_num);
				maxrow=sheet.getLastRowNum();  //获得该sheet工作表上最后有值的行
				wb.close();
			}
			else                        //.xlsx后缀        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel的sheet
				sheet = xwb.getSheetAt(sheet_num);
				maxrow=sheet.getLastRowNum();  //获得该sheet工作表上最后有值的行
				xwb.close();
			}
	    }
		catch (IOException e)
		{
			e.printStackTrace();
		}
		return maxrow+1;

	}
	
	/**	
	    * 获得一个EXCEL工作簿指定工作表的最大行数,通过工作表名称打开<br>
	    * 如返回值=4(0、1、2、3),则第4行及后面的行上的单元格都没有内容
	    * @param File_Path
	    *        文件的绝对路径
	    * @param sheet_name
	    * 		  工作簿中的工作表名称
	    * @return int
	*/
	public static int getmaxrow(String File_Path,String sheet_name)
	{
		int maxrow=0;
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls后缀           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel的sheet
				sheet = wb.getSheet(sheet_name);
				maxrow=sheet.getLastRowNum();  //获得该sheet工作表上最后有值的行
				wb.close();
			}
			else                        //.xlsx后缀        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel的sheet
				sheet = xwb.getSheet(sheet_name);
				maxrow=sheet.getLastRowNum();  //获得该sheet工作表上最后有值的行
				xwb.close();
			}
	    }
		catch (IOException e)
		{
			e.printStackTrace();
		}
		return maxrow+1;

	}
	
	/**	
	    * 获得一个EXCEL工作簿指定工作表上指定行的最大列数<br>
	    * 如hang=0该行上,B列是不为空的单元格(C、D等后面的列单元格都是空)，返回2
	    * @param File_Path
	    *        文件的绝对路径
	    * @param sheet_num
	    * 		 工作簿中的工作表序号,第一个工作表序号为0
	    * @param hang
	    * 		 工作表的行数,第一行为0
	    * @return int
	*/
	public static int getmaxcell(String File_Path,int sheet_num,int hang)
	{
		int maxcell=0;
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls后缀           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel的sheet
				HSSFRow row;             //excel的行
				
				sheet = wb.getSheetAt(sheet_num);
				row = sheet.getRow(hang);
				maxcell=row.getLastCellNum();  //获得该sheet工作表上最后一个不为空的单元格列
				
				wb.close();
			}
			else                        //.xlsx后缀        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel的sheet
				XSSFRow row;             //excel的行
				
				sheet = xwb.getSheetAt(sheet_num);
				row = sheet.getRow(hang);
				maxcell=row.getLastCellNum();  //获得该sheet工作表上最后一个不为空的单元格列
				
				xwb.close();
			}
			
	    }
		catch (IOException e)
		{
			e.printStackTrace();
		}
		return maxcell+1;

	}
	
	/**	
	    * read_wholecell：读取EXCEL整列内容,通过序号打开工作表<br>
	    * @param File_Path
	    *        文件绝对路径
	    * @param sheet_num
	    * 		 工作簿中的工作表序号,第一个工作表序号为0
	    * @param lie
	    * 		 读取的列,第一列为0
	*/
	public static void read_wholecell(String File_Path, int sheet_num, int lie, ArrayList <String> Arraylist)
	{
		int hang;
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls后缀           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel的sheet
				HSSFRow row;             //excel的行
				HSSFCell cell;           //excel的列
				
				sheet = wb.getSheetAt(sheet_num);
				
				//合并单元格系列
				//start
				ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
				int sheetmergerCount = sheet.getNumMergedRegions();
				CellRangeAddress ca;
				if(sheetmergerCount !=0 )
				{
					for (int i=0; i < sheetmergerCount; i++)
					{
						  // 获得合并单元格加入list中
						  ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //返回指定索引合并后的区域
						  LIST.add(ca);
					}
				}
				//end
				
				int maxrow=sheet.getLastRowNum();  //获得该sheet工作表上最后有文字行数
				for(hang=0; hang<maxrow; hang++){
					if(hang>maxrow){
						//System.out.println("超过最大行数);
						Arraylist.add("");
						continue;
					}
					row = sheet.getRow(hang);
					if(row==null){
						Arraylist.add("");
						continue;
					}
					cell=row.getCell(lie);
					if(cell==null || cell.toString()==(""))
					{
						//System.out.println("读到空");
						//判断是否属于合并单元格
						if(sheetmergerCount==0)
						{
							Arraylist.add("");
							continue;
						}
						
						int firstC = 0;
						int lastC = 0;
						int firstR = 0;
						int lastR = 0;
						
						int run_i;
						for (run_i=0;run_i<LIST.size();run_i++)
						{
							ca=LIST.get(run_i);
							// 获得合并单元格的起始行, 结束行, 起始列, 结束列
							firstC = ca.getFirstColumn();
							lastC = ca.getLastColumn();
							firstR = ca.getFirstRow();
							lastR = ca.getLastRow();
							if (lie <= lastC && lie>= firstC)
							{
								if (hang <= lastR && hang >= firstR)
								{
									//合并单元格的值，读取合并区域的首行首列
									row = sheet.getRow(firstR);
									cell=row.getCell(firstC);
									break;
								}
							}
						}
						
						//不是合并单元格,其单元格内没有内容
						if(run_i >= LIST.size())
						{
							Arraylist.add("");
							continue;
						}
						
						//合并区域的首行首列无内容
						if(cell==null || cell.toString()==(""))
						{
							Arraylist.add("");
							continue;
						}
					}
					switch (cell.getCellType()) 
					{
						case HSSFCell.CELL_TYPE_STRING: //字符串类型
							Arraylist.add( cell.getStringCellValue() );
							//str=cell.getStringCellValue();
							break;
						case HSSFCell.CELL_TYPE_NUMERIC: //数值类型
							double qudian;
							qudian=cell.getNumericCellValue();
							Arraylist.add( String.format("%.0f", qudian) );
							//str=String.format("%.0f", qudian);
							//String.valueOf(hex.getNumericCellValue()); 
							break;
					}
				}
				wb.close();
				
			}
			else         //.xlsx后缀        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel的sheet
				XSSFRow row;             //excel的行
				XSSFCell cell;           //excel的列
				
				sheet = xwb.getSheetAt(sheet_num);
				
				//合并单元格系列
				//start
				ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
				int sheetmergerCount = sheet.getNumMergedRegions();
				CellRangeAddress ca;
				if(sheetmergerCount !=0 )
				{
					for (int i=0; i < sheetmergerCount; i++)
					{
						  // 获得合并单元格加入list中
						  ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //返回指定索引合并后的区域
						  LIST.add(ca);
					}
				}
				//end
				
				int maxrow=sheet.getLastRowNum();  //获得该sheet工作表上最后有文字行数
				for(hang=0; hang<maxrow; hang++){
					if(hang>maxrow){
						//System.out.println("超过最大行数);
						Arraylist.add("");
						continue;
					}
					row = sheet.getRow(hang);
					if(row==null){
						Arraylist.add("");
						continue;
					}
					cell=row.getCell(lie);
					if(cell==null || cell.toString()==(""))
					{
						//System.out.println("读到空");
						//判断是否属于合并单元格
						if(sheetmergerCount==0)
						{
							Arraylist.add("");
							continue;
						}
						int firstC = 0;
						int lastC = 0;
						int firstR = 0;
						int lastR = 0;
						
						int run_i;
						for (run_i=0;run_i<LIST.size();run_i++)
						{
							ca=LIST.get(run_i);
							// 获得合并单元格的起始行, 结束行, 起始列, 结束列
							firstC = ca.getFirstColumn();
							lastC = ca.getLastColumn();
							firstR = ca.getFirstRow();
							lastR = ca.getLastRow();
							if (lie <= lastC && lie>= firstC)
							{
								if (hang <= lastR && hang >= firstR)
								{
									//合并单元格的值，读取合并区域的首行首列
									row = sheet.getRow(firstR);
									cell=row.getCell(firstC);
									break;
								}
							}
						}
						//不是合并单元格,其单元格内没有内容
						if(run_i>=LIST.size())
						{
							Arraylist.add("");
							continue;
						}
						//合并区域的首行首列无内容
						if(cell==null || cell.toString()==(""))
						{
							Arraylist.add("");
							continue;
						}
					}
					switch (cell.getCellType()) 
					{
						case HSSFCell.CELL_TYPE_STRING: //字符串类型
							Arraylist.add( cell.getStringCellValue() );
							//str=cell.getStringCellValue();
							break;
						case HSSFCell.CELL_TYPE_NUMERIC: //数值类型
							double qudian;
							qudian=cell.getNumericCellValue();
							Arraylist.add( String.format("%.0f", qudian) );
							//str=String.format("%.0f", qudian);
							//String.valueOf(hex.getNumericCellValue()); 
							break;
					}
				}
				xwb.close();
			}
	    }
		catch (IOException e)
		{
			e.printStackTrace();
		}
	}

	/**	
	    * read_wholecell：读取EXCEL整列内容,通过名称打开工作表<br>
	    * @param File_Path
	    *        文件绝对路径
	    * @param sheet_name
	    * 		 工作簿中的工作表名称
	    * @param lie
	    * 		 读取的列,第一列为0
	*/
	public static void read_wholecell(String File_Path, String sheet_name, int lie, ArrayList <String> Arraylist)
	{
		int hang;
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls后缀           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel的sheet
				HSSFRow row;             //excel的行
				HSSFCell cell;           //excel的列
				
				sheet = wb.getSheet(sheet_name);
				
				//合并单元格系列
				//start
				ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
				int sheetmergerCount = sheet.getNumMergedRegions();
				CellRangeAddress ca;
				if(sheetmergerCount !=0 )
				{
					for (int i=0; i < sheetmergerCount; i++)
					{
						  // 获得合并单元格加入list中
						  ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //返回指定索引合并后的区域
						  LIST.add(ca);
					}
				}
				//end
				
				int maxrow=sheet.getLastRowNum();  //获得该sheet工作表上最后有文字行数
				for(hang=0; hang<maxrow; hang++){
					if(hang>maxrow){
						//System.out.println("超过最大行数);
						Arraylist.add("");
						continue;
					}
					row = sheet.getRow(hang);
					if(row==null){
						Arraylist.add("");
						continue;
					}
					cell=row.getCell(lie);
					if(cell==null || cell.toString()==(""))
					{
						//System.out.println("读到空");
						//判断是否属于合并单元格
						if(sheetmergerCount==0)
						{
							Arraylist.add("");
							continue;
						}
						
						int firstC = 0;
						int lastC = 0;
						int firstR = 0;
						int lastR = 0;
						
						int run_i;
						for (run_i=0;run_i<LIST.size();run_i++)
						{
							ca=LIST.get(run_i);
							// 获得合并单元格的起始行, 结束行, 起始列, 结束列
							firstC = ca.getFirstColumn();
							lastC = ca.getLastColumn();
							firstR = ca.getFirstRow();
							lastR = ca.getLastRow();
							if (lie <= lastC && lie>= firstC)
							{
								if (hang <= lastR && hang >= firstR)
								{
									//合并单元格的值，读取合并区域的首行首列
									row = sheet.getRow(firstR);
									cell=row.getCell(firstC);
									break;
								}
							}
						}
						
						//不是合并单元格,其单元格内没有内容
						if(run_i>=LIST.size())
						{
							Arraylist.add("");
							continue;
						}
						
						//合并区域的首行首列无内容
						if(cell==null || cell.toString()==(""))
						{
							Arraylist.add("");
							continue;
						}
					}
					switch (cell.getCellType()) 
					{
						case HSSFCell.CELL_TYPE_STRING: //字符串类型
							Arraylist.add( cell.getStringCellValue() );
							//str=cell.getStringCellValue();
							break;
						case HSSFCell.CELL_TYPE_NUMERIC: //数值类型
							double qudian;
							qudian=cell.getNumericCellValue();
							Arraylist.add( String.format("%.0f", qudian) );
							//str=String.format("%.0f", qudian);
							//String.valueOf(hex.getNumericCellValue()); 
							break;
					}
				}
				wb.close();
				
			}
			else         //.xlsx后缀        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel的sheet
				XSSFRow row;             //excel的行
				XSSFCell cell;           //excel的列
				
				sheet = xwb.getSheet(sheet_name);
				
				//合并单元格系列
				//start
				ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
				int sheetmergerCount = sheet.getNumMergedRegions();
				CellRangeAddress ca;
				if(sheetmergerCount !=0 )
				{
					for (int i=0; i < sheetmergerCount; i++)
					{
						  // 获得合并单元格加入list中
						  ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //返回指定索引合并后的区域
						  LIST.add(ca);
					}
				}
				//end
				
				int maxrow=sheet.getLastRowNum();  //获得该sheet工作表上最后有文字行数
				for(hang=0; hang<maxrow; hang++){
					if(hang>maxrow){
						//System.out.println("超过最大行数);
						Arraylist.add("");
						continue;
					}
					row = sheet.getRow(hang);
					if(row==null){
						Arraylist.add("");
						continue;
					}
					cell=row.getCell(lie);
					if(cell==null || cell.toString()==(""))
					{
						//System.out.println("读到空");
						//判断是否属于合并单元格
						if(sheetmergerCount==0)
						{
							Arraylist.add("");
							continue;
						}
						int firstC = 0;
						int lastC = 0;
						int firstR = 0;
						int lastR = 0;
						
						int run_i;
						for (run_i=0;run_i<LIST.size();run_i++)
						{
							ca=LIST.get(run_i);
							// 获得合并单元格的起始行, 结束行, 起始列, 结束列
							firstC = ca.getFirstColumn();
							lastC = ca.getLastColumn();
							firstR = ca.getFirstRow();
							lastR = ca.getLastRow();
							if (lie <= lastC && lie>= firstC)
							{
								if (hang <= lastR && hang >= firstR)
								{
									//合并单元格的值，读取合并区域的首行首列
									row = sheet.getRow(firstR);
									cell=row.getCell(firstC);
									break;
								}
							}
						}
						//不是合并单元格,其单元格内没有内容
						if(run_i>=LIST.size())
						{
							Arraylist.add("");
							continue;
						}
						//合并区域的首行首列无内容
						if(cell==null || cell.toString()==(""))
						{
							Arraylist.add("");
							continue;
						}
					}
					switch (cell.getCellType()) 
					{
						case HSSFCell.CELL_TYPE_STRING: //字符串类型
							Arraylist.add( cell.getStringCellValue() );
							//str=cell.getStringCellValue();
							break;
						case HSSFCell.CELL_TYPE_NUMERIC: //数值类型
							double qudian;
							qudian=cell.getNumericCellValue();
							Arraylist.add( String.format("%.0f", qudian) );
							//str=String.format("%.0f", qudian);
							//String.valueOf(hex.getNumericCellValue()); 
							break;
					}
				}
				xwb.close();
			}
	    }
		catch (IOException e)
		{
			e.printStackTrace();
		}
	}
	
	
	

}
