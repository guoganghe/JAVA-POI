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
	    * ͨ�����������Ŵ򿪹�����,�ܶ�ȡ�ϲ���Ԫ���ֵ
	    * @param File_Path
	    *        �ļ��ľ���·��
	    * @param sheet_num
	    * 		 �������еĹ�������ţ���һ���������Ϊ0
	    * @param hang
	    * 		 �������е�����,EXCEL�ļ��еĵ�1�м�Ϊ0
	    * @param lie
	    * 		 �������е�����,EXCEL�ļ��еĵ�A�м�Ϊ0
	    * @return String
	*/
	public static String read(String File_Path,int sheet_num,int hang,int lie)
	{
		String str="";
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls��׺           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel��sheet
				HSSFRow row;             //excel����
				HSSFCell cell;           //excel����
				
				sheet = wb.getSheetAt(sheet_num);
				int maxrow=sheet.getLastRowNum();  //��ø�sheet�����������������
				if(hang>maxrow)
				{
					//System.out.println("���������Ч��������:"+hang+",�������Ϊ:"+maxrow);
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
				int maxcell=row.getLastCellNum();  //��ø�sheet�������������
				if(lie>maxcell)
				{
					//System.out.println(hang+"��--"+"���������Ч��������:"+lie+",�������Ϊ:"+maxcell);
					wb.close();
					return null;
				}*/
				cell=row.getCell(lie);

				if(cell==null || cell.toString()==(""))
				{
					//System.out.println("������");
					
					ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
					//�ж��Ƿ��Ǻϲ���Ԫ��
					int sheetmergerCount = sheet.getNumMergedRegions();
					if(sheetmergerCount==0)
					{
						wb.close();
						return null;
					}
					
					for (int i = 0; i < sheetmergerCount; i++)
					{
						  // ��úϲ���Ԫ�����list��
						  CellRangeAddress ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //����ָ�������ϲ��������
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
						// ��úϲ���Ԫ�����ʼ��, ������, ��ʼ��, ������
						firstC = ca.getFirstColumn();
						lastC = ca.getLastColumn();
						firstR = ca.getFirstRow();
						lastR = ca.getLastRow();
						if (lie <= lastC&& lie>= firstC)
						{
							if (hang <= lastR && hang >= firstR)
							{
								//�ϲ���Ԫ���ֵ����ȡ�ϲ��������������
								row = sheet.getRow(firstR);
								cell=row.getCell(firstC);
								break;
							}
						}
					}
					
					//���Ǻϲ���Ԫ��,�䵥Ԫ����û������
					if(run_i>=LIST.size())
					{
						wb.close();
						return null;
					}
					
					//�ϲ��������������������
					if(cell==null || cell.toString()==(""))
					{
						wb.close();
						return null;
					}
				}
				switch (cell.getCellType()) 
				{
					case HSSFCell.CELL_TYPE_STRING: //�ַ�������
						str=cell.getStringCellValue();
						break;
					case HSSFCell.CELL_TYPE_NUMERIC: //��ֵ����
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
			else         //.xlsx��׺        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel��sheet
				XSSFRow row;             //excel����
				XSSFCell cell;           //excel����
				
				sheet = xwb.getSheetAt(sheet_num);
				
				int maxrow=sheet.getLastRowNum();  //��ø�sheet�����������������
				if(hang>maxrow)
				{
					//System.out.println("���������Ч��������:"+hang+",�������Ϊ:"+maxrow);
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
				int maxcell=row.getLastCellNum();  //��ø�sheet�����������������
				if(lie>maxcell)
				{
					//System.out.println(hang+"��--"+"���������Ч��������:"+lie+",�������Ϊ:"+maxcell);
					xwb.close();
					return null;
				}*/
				
				cell=row.getCell(lie);
				if(cell==null || cell.toString()==(""))
				{
					//System.out.println("������");
					
					ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
					//�ж��Ƿ��Ǻϲ���Ԫ��
					int sheetmergerCount = sheet.getNumMergedRegions();
					if(sheetmergerCount==0)
					{
						xwb.close();
						return null;
					}
					
					for (int i = 0; i < sheetmergerCount; i++)
					{
						  // ��úϲ���Ԫ�����list��
						  CellRangeAddress ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //����ָ�������ϲ��������
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
						// ��úϲ���Ԫ�����ʼ��, ������, ��ʼ��, ������
						firstC = ca.getFirstColumn();
						lastC = ca.getLastColumn();
						firstR = ca.getFirstRow();
						lastR = ca.getLastRow();
						if (lie <= lastC&& lie>= firstC)
						{
							if (hang <= lastR && hang >= firstR)
							{
								//�ϲ���Ԫ���ֵ����ȡ�ϲ��������������
								row = sheet.getRow(firstR);
								cell=row.getCell(firstC);
								break;
							}
						}
					}
					
					//���Ǻϲ���Ԫ��,�䵥Ԫ����û������
					if(run_i>=LIST.size())
					{
						xwb.close();
						return null;
					}
					
					//�ϲ��������������������
					if(cell==null || cell.toString()==(""))
					{
						xwb.close();
						return null;
					}
				}
				switch (cell.getCellType()) 
				{
					case XSSFCell.CELL_TYPE_STRING: //�ַ�������
						str=cell.getStringCellValue();
						break;
					case XSSFCell.CELL_TYPE_NUMERIC: //��ֵ����
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
	    * ͨ������������ƴ򿪹�����,�ܶ�ȡ�ϲ���Ԫ���ֵ
	    * @param File_Path
	    *        �ļ��ľ���·��
	    * @param sheet_name
	    * 		 �������еĹ���������
	    * @param hang
	    * 		 �������е�����,EXCEL�ļ��еĵ�1�м�Ϊ0
	    * @param lie
	    * 		 �������е�����,EXCEL�ļ��еĵ�A�м�Ϊ0
	    * @return String
	*/
	public static String read(String File_Path,String sheet_name,int hang,int lie)
	{
		String str="";
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls��׺           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel��sheet
				HSSFRow row;             //excel����
				HSSFCell cell;           //excel����
				
				sheet = wb.getSheet(sheet_name);
				
				int maxrow=sheet.getLastRowNum();  //��ø�sheet�����������������
				if(hang>maxrow)
				{
					//System.out.println("���������Ч��������:"+hang+",�������Ϊ:"+maxrow);
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
				int maxcell=row.getLastCellNum();  //��ø�sheet�����������������
				if(lie>maxcell)
				{
					//System.out.println(hang+"��--"+"���������Ч��������:"+lie+",�������Ϊ:"+maxcell);
					wb.close();
					return null;
				}*/
				cell=row.getCell(lie);
				
				if(cell==null || cell.toString()==(""))
				{
					//System.out.println("������");
					
					ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
					//�ж��Ƿ��Ǻϲ���Ԫ��
					int sheetmergerCount = sheet.getNumMergedRegions();
					if(sheetmergerCount==0)
					{
						wb.close();
						return null;
					}
					
					for (int i = 0; i < sheetmergerCount; i++)
					{
						  // ��úϲ���Ԫ�����list��
						  CellRangeAddress ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //����ָ�������ϲ��������
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
						// ��úϲ���Ԫ�����ʼ��, ������, ��ʼ��, ������
						firstC = ca.getFirstColumn();
						lastC = ca.getLastColumn();
						firstR = ca.getFirstRow();
						lastR = ca.getLastRow();
						if (lie <= lastC&& lie>= firstC)
						{
							if (hang <= lastR && hang >= firstR)
							{
								//�ϲ���Ԫ���ֵ����ȡ�ϲ��������������
								row = sheet.getRow(firstR);
								cell=row.getCell(firstC);
								break;
							}
						}
					}
					
					//���Ǻϲ���Ԫ��,�䵥Ԫ����û������
					if(run_i>=LIST.size())
					{
						wb.close();
						return null;
					}
					
					//�ϲ��������������������
					if(cell==null || cell.toString()==(""))
					{
						wb.close();
						return null;
					}
				}
				switch (cell.getCellType()) 
				{
					case HSSFCell.CELL_TYPE_STRING: //�ַ�������
						str=cell.getStringCellValue();
						break;
					case HSSFCell.CELL_TYPE_NUMERIC: //��ֵ����
						double qudian;
						qudian=cell.getNumericCellValue();
						str=String.format("%.0f", qudian);
						//codehex=String.valueOf(hex.getNumericCellValue()); 
						break;
				}
				wb.close();
				
			}
			else         //.xlsx��׺        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel��sheet
				XSSFRow row;             //excel����
				XSSFCell cell;           //excel����
				
				sheet = xwb.getSheet(sheet_name);
				
				int maxrow=sheet.getLastRowNum();  //��ø�sheet�����������������
				if(hang>maxrow)
				{
					//System.out.println("���������Ч��������:"+hang+",�������Ϊ:"+maxrow);
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
				int maxcell=row.getLastCellNum();  //��ø�sheet�����������������
				if(lie>maxcell)
				{
					//System.out.println(hang+"��--"+"���������Ч��������:"+lie+",�������Ϊ:"+maxcell);
					xwb.close();
					return null;
				}*/
				
				cell=row.getCell(lie);
				if(cell==null || cell.toString()==(""))
				{
					//System.out.println("������");
					
					ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
					//�ж��Ƿ��Ǻϲ���Ԫ��
					int sheetmergerCount = sheet.getNumMergedRegions();
					if(sheetmergerCount==0)
					{
						xwb.close();
						return null;
					}
					
					for (int i = 0; i < sheetmergerCount; i++)
					{
						  // ��úϲ���Ԫ�����list��
						  CellRangeAddress ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //����ָ�������ϲ��������
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
						// ��úϲ���Ԫ�����ʼ��, ������, ��ʼ��, ������
						firstC = ca.getFirstColumn();
						lastC = ca.getLastColumn();
						firstR = ca.getFirstRow();
						lastR = ca.getLastRow();
						if (lie <= lastC&& lie>= firstC)
						{
							if (hang <= lastR && hang >= firstR)
							{
								//�ϲ���Ԫ���ֵ����ȡ�ϲ��������������
								row = sheet.getRow(firstR);
								cell=row.getCell(firstC);
								break;
							}
						}
					}
					
					//���Ǻϲ���Ԫ��,�䵥Ԫ����û������
					if(run_i>=LIST.size())
					{
						xwb.close();
						return null;
					}
					
					//�ϲ��������������������
					if(cell==null || cell.toString()==(""))
					{
						xwb.close();
						return null;
					}
				}
				switch (cell.getCellType()) 
				{
					case XSSFCell.CELL_TYPE_STRING: //�ַ�������
						str=cell.getStringCellValue();
						break;
					case XSSFCell.CELL_TYPE_NUMERIC: //��ֵ����
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
	    * ���һ��EXCEL�ļ�ָ����ŵĹ���������<br>
	    * ��һ�������������0
	    * @param File_Path
	    *        �ļ��ľ���·��
	    * @param sheet_num
	    * 		 �������еĹ��������,��һ����������ż�Ϊ0
	    * @return String
	*/
	public static String getsheetname(String File_Path,int sheet_num)
	{
		String sheetname="";
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls��׺           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				sheetname=wb.getSheetName(sheet_num);  //���sheet����
				wb.close();
			}
			else                        //.xlsx��׺        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				sheetname=xwb.getSheetName(sheet_num);  //���sheet����
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
	    * ���һ��EXCEL�������Ĺ���������<br>
	    * �緵��ֵ=2,����2�������������Ϊ0��1
	    * @param File_Path
	    *        �ļ��ľ���·��
	    * @return int
	*/
	public static int getsheetnum(String File_Path)
	{
		int sheet_num=0;
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls��׺           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				sheet_num=wb.getNumberOfSheets();  //���sheet����
				wb.close();
			}
			else                        //.xlsx��׺        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				sheet_num=xwb.getNumberOfSheets();  //���sheet����
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
	    * ���һ��EXCEL������ָ����������������,ͨ����������Ŵ�<br>
	    * �緵��ֵ=4(0��1��2��3),���4�м���������ϵĵ�Ԫ��û������
	    * @param File_Path
	    *        �ļ��ľ���·��
	    * @param sheet_num
	    * 		  �������еĹ��������,��һ����������ż�Ϊ0
	    * @return int
	*/
	public static int getmaxrow(String File_Path,int sheet_num)
	{
		int maxrow=0;
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls��׺           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel��sheet
				sheet = wb.getSheetAt(sheet_num);
				maxrow=sheet.getLastRowNum();  //��ø�sheet�������������ֵ����
				wb.close();
			}
			else                        //.xlsx��׺        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel��sheet
				sheet = xwb.getSheetAt(sheet_num);
				maxrow=sheet.getLastRowNum();  //��ø�sheet�������������ֵ����
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
	    * ���һ��EXCEL������ָ����������������,ͨ�����������ƴ�<br>
	    * �緵��ֵ=4(0��1��2��3),���4�м���������ϵĵ�Ԫ��û������
	    * @param File_Path
	    *        �ļ��ľ���·��
	    * @param sheet_name
	    * 		  �������еĹ���������
	    * @return int
	*/
	public static int getmaxrow(String File_Path,String sheet_name)
	{
		int maxrow=0;
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls��׺           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel��sheet
				sheet = wb.getSheet(sheet_name);
				maxrow=sheet.getLastRowNum();  //��ø�sheet�������������ֵ����
				wb.close();
			}
			else                        //.xlsx��׺        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel��sheet
				sheet = xwb.getSheet(sheet_name);
				maxrow=sheet.getLastRowNum();  //��ø�sheet�������������ֵ����
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
	    * ���һ��EXCEL������ָ����������ָ���е��������<br>
	    * ��hang=0������,B���ǲ�Ϊ�յĵ�Ԫ��(C��D�Ⱥ�����е�Ԫ���ǿ�)������2
	    * @param File_Path
	    *        �ļ��ľ���·��
	    * @param sheet_num
	    * 		 �������еĹ��������,��һ�����������Ϊ0
	    * @param hang
	    * 		 �����������,��һ��Ϊ0
	    * @return int
	*/
	public static int getmaxcell(String File_Path,int sheet_num,int hang)
	{
		int maxcell=0;
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls��׺           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel��sheet
				HSSFRow row;             //excel����
				
				sheet = wb.getSheetAt(sheet_num);
				row = sheet.getRow(hang);
				maxcell=row.getLastCellNum();  //��ø�sheet�����������һ����Ϊ�յĵ�Ԫ����
				
				wb.close();
			}
			else                        //.xlsx��׺        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel��sheet
				XSSFRow row;             //excel����
				
				sheet = xwb.getSheetAt(sheet_num);
				row = sheet.getRow(hang);
				maxcell=row.getLastCellNum();  //��ø�sheet�����������һ����Ϊ�յĵ�Ԫ����
				
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
	    * read_wholecell����ȡEXCEL��������,ͨ����Ŵ򿪹�����<br>
	    * @param File_Path
	    *        �ļ�����·��
	    * @param sheet_num
	    * 		 �������еĹ��������,��һ�����������Ϊ0
	    * @param lie
	    * 		 ��ȡ����,��һ��Ϊ0
	*/
	public static void read_wholecell(String File_Path, int sheet_num, int lie, ArrayList <String> Arraylist)
	{
		int hang;
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls��׺           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel��sheet
				HSSFRow row;             //excel����
				HSSFCell cell;           //excel����
				
				sheet = wb.getSheetAt(sheet_num);
				
				//�ϲ���Ԫ��ϵ��
				//start
				ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
				int sheetmergerCount = sheet.getNumMergedRegions();
				CellRangeAddress ca;
				if(sheetmergerCount !=0 )
				{
					for (int i=0; i < sheetmergerCount; i++)
					{
						  // ��úϲ���Ԫ�����list��
						  ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //����ָ�������ϲ��������
						  LIST.add(ca);
					}
				}
				//end
				
				int maxrow=sheet.getLastRowNum();  //��ø�sheet���������������������
				for(hang=0; hang<maxrow; hang++){
					if(hang>maxrow){
						//System.out.println("�����������);
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
						//System.out.println("������");
						//�ж��Ƿ����ںϲ���Ԫ��
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
							// ��úϲ���Ԫ�����ʼ��, ������, ��ʼ��, ������
							firstC = ca.getFirstColumn();
							lastC = ca.getLastColumn();
							firstR = ca.getFirstRow();
							lastR = ca.getLastRow();
							if (lie <= lastC && lie>= firstC)
							{
								if (hang <= lastR && hang >= firstR)
								{
									//�ϲ���Ԫ���ֵ����ȡ�ϲ��������������
									row = sheet.getRow(firstR);
									cell=row.getCell(firstC);
									break;
								}
							}
						}
						
						//���Ǻϲ���Ԫ��,�䵥Ԫ����û������
						if(run_i >= LIST.size())
						{
							Arraylist.add("");
							continue;
						}
						
						//�ϲ��������������������
						if(cell==null || cell.toString()==(""))
						{
							Arraylist.add("");
							continue;
						}
					}
					switch (cell.getCellType()) 
					{
						case HSSFCell.CELL_TYPE_STRING: //�ַ�������
							Arraylist.add( cell.getStringCellValue() );
							//str=cell.getStringCellValue();
							break;
						case HSSFCell.CELL_TYPE_NUMERIC: //��ֵ����
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
			else         //.xlsx��׺        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel��sheet
				XSSFRow row;             //excel����
				XSSFCell cell;           //excel����
				
				sheet = xwb.getSheetAt(sheet_num);
				
				//�ϲ���Ԫ��ϵ��
				//start
				ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
				int sheetmergerCount = sheet.getNumMergedRegions();
				CellRangeAddress ca;
				if(sheetmergerCount !=0 )
				{
					for (int i=0; i < sheetmergerCount; i++)
					{
						  // ��úϲ���Ԫ�����list��
						  ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //����ָ�������ϲ��������
						  LIST.add(ca);
					}
				}
				//end
				
				int maxrow=sheet.getLastRowNum();  //��ø�sheet���������������������
				for(hang=0; hang<maxrow; hang++){
					if(hang>maxrow){
						//System.out.println("�����������);
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
						//System.out.println("������");
						//�ж��Ƿ����ںϲ���Ԫ��
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
							// ��úϲ���Ԫ�����ʼ��, ������, ��ʼ��, ������
							firstC = ca.getFirstColumn();
							lastC = ca.getLastColumn();
							firstR = ca.getFirstRow();
							lastR = ca.getLastRow();
							if (lie <= lastC && lie>= firstC)
							{
								if (hang <= lastR && hang >= firstR)
								{
									//�ϲ���Ԫ���ֵ����ȡ�ϲ��������������
									row = sheet.getRow(firstR);
									cell=row.getCell(firstC);
									break;
								}
							}
						}
						//���Ǻϲ���Ԫ��,�䵥Ԫ����û������
						if(run_i>=LIST.size())
						{
							Arraylist.add("");
							continue;
						}
						//�ϲ��������������������
						if(cell==null || cell.toString()==(""))
						{
							Arraylist.add("");
							continue;
						}
					}
					switch (cell.getCellType()) 
					{
						case HSSFCell.CELL_TYPE_STRING: //�ַ�������
							Arraylist.add( cell.getStringCellValue() );
							//str=cell.getStringCellValue();
							break;
						case HSSFCell.CELL_TYPE_NUMERIC: //��ֵ����
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
	    * read_wholecell����ȡEXCEL��������,ͨ�����ƴ򿪹�����<br>
	    * @param File_Path
	    *        �ļ�����·��
	    * @param sheet_name
	    * 		 �������еĹ���������
	    * @param lie
	    * 		 ��ȡ����,��һ��Ϊ0
	*/
	public static void read_wholecell(String File_Path, String sheet_name, int lie, ArrayList <String> Arraylist)
	{
		int hang;
		try
	    {
			int length=File_Path.length();
			String buff=File_Path.substring(length-4);
			if(buff.equals(".xls"))     //.xls��׺           EXCEL2003
			{
				HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(File_Path)); 
				HSSFSheet sheet;         //excel��sheet
				HSSFRow row;             //excel����
				HSSFCell cell;           //excel����
				
				sheet = wb.getSheet(sheet_name);
				
				//�ϲ���Ԫ��ϵ��
				//start
				ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
				int sheetmergerCount = sheet.getNumMergedRegions();
				CellRangeAddress ca;
				if(sheetmergerCount !=0 )
				{
					for (int i=0; i < sheetmergerCount; i++)
					{
						  // ��úϲ���Ԫ�����list��
						  ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //����ָ�������ϲ��������
						  LIST.add(ca);
					}
				}
				//end
				
				int maxrow=sheet.getLastRowNum();  //��ø�sheet���������������������
				for(hang=0; hang<maxrow; hang++){
					if(hang>maxrow){
						//System.out.println("�����������);
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
						//System.out.println("������");
						//�ж��Ƿ����ںϲ���Ԫ��
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
							// ��úϲ���Ԫ�����ʼ��, ������, ��ʼ��, ������
							firstC = ca.getFirstColumn();
							lastC = ca.getLastColumn();
							firstR = ca.getFirstRow();
							lastR = ca.getLastRow();
							if (lie <= lastC && lie>= firstC)
							{
								if (hang <= lastR && hang >= firstR)
								{
									//�ϲ���Ԫ���ֵ����ȡ�ϲ��������������
									row = sheet.getRow(firstR);
									cell=row.getCell(firstC);
									break;
								}
							}
						}
						
						//���Ǻϲ���Ԫ��,�䵥Ԫ����û������
						if(run_i>=LIST.size())
						{
							Arraylist.add("");
							continue;
						}
						
						//�ϲ��������������������
						if(cell==null || cell.toString()==(""))
						{
							Arraylist.add("");
							continue;
						}
					}
					switch (cell.getCellType()) 
					{
						case HSSFCell.CELL_TYPE_STRING: //�ַ�������
							Arraylist.add( cell.getStringCellValue() );
							//str=cell.getStringCellValue();
							break;
						case HSSFCell.CELL_TYPE_NUMERIC: //��ֵ����
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
			else         //.xlsx��׺        EXCEL2007
			{
				XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(File_Path)); 
				XSSFSheet sheet;         //excel��sheet
				XSSFRow row;             //excel����
				XSSFCell cell;           //excel����
				
				sheet = xwb.getSheet(sheet_name);
				
				//�ϲ���Ԫ��ϵ��
				//start
				ArrayList <CellRangeAddress> LIST=new ArrayList <CellRangeAddress>();
				int sheetmergerCount = sheet.getNumMergedRegions();
				CellRangeAddress ca;
				if(sheetmergerCount !=0 )
				{
					for (int i=0; i < sheetmergerCount; i++)
					{
						  // ��úϲ���Ԫ�����list��
						  ca = sheet.getMergedRegion(i);
						  //Returns the merged region at the specified index
						  //����ָ�������ϲ��������
						  LIST.add(ca);
					}
				}
				//end
				
				int maxrow=sheet.getLastRowNum();  //��ø�sheet���������������������
				for(hang=0; hang<maxrow; hang++){
					if(hang>maxrow){
						//System.out.println("�����������);
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
						//System.out.println("������");
						//�ж��Ƿ����ںϲ���Ԫ��
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
							// ��úϲ���Ԫ�����ʼ��, ������, ��ʼ��, ������
							firstC = ca.getFirstColumn();
							lastC = ca.getLastColumn();
							firstR = ca.getFirstRow();
							lastR = ca.getLastRow();
							if (lie <= lastC && lie>= firstC)
							{
								if (hang <= lastR && hang >= firstR)
								{
									//�ϲ���Ԫ���ֵ����ȡ�ϲ��������������
									row = sheet.getRow(firstR);
									cell=row.getCell(firstC);
									break;
								}
							}
						}
						//���Ǻϲ���Ԫ��,�䵥Ԫ����û������
						if(run_i>=LIST.size())
						{
							Arraylist.add("");
							continue;
						}
						//�ϲ��������������������
						if(cell==null || cell.toString()==(""))
						{
							Arraylist.add("");
							continue;
						}
					}
					switch (cell.getCellType()) 
					{
						case HSSFCell.CELL_TYPE_STRING: //�ַ�������
							Arraylist.add( cell.getStringCellValue() );
							//str=cell.getStringCellValue();
							break;
						case HSSFCell.CELL_TYPE_NUMERIC: //��ֵ����
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
