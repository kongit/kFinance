package k.finance;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import org.apache.commons.lang3.StringUtils;

public class GetContentFromXls {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		String file = "d://cuc.xls";
		GetContentFromXls.get(file);
	}
	
	public static void get(String file){
		Workbook workbook = null;
		try {
			workbook = Workbook.getWorkbook(new File(file));
			Sheet sheet = workbook.getSheet(0);
			//获得文件中一共有多少行
			int totalRows = sheet.getRows();
			System.out.println("总行数：" + totalRows);
			
			//数据处理结果list
			List<Map<String,String>> okList = new ArrayList<Map<String,String>>();
			
			List<Map<String,Integer>> groupList = getGroupList(sheet, totalRows);
			//得到每组的行数据
			for(Map<String,Integer> groupMap : groupList){
				
				
				
				int sIdx = groupMap.get("s");
				int eIdx = groupMap.get("e");
				//获得项目编号，在每组的第一行。用. split 之后 第5位是项目编号
				String okProjCode = StringUtils.split(getCellValue2Trim(sheet, sIdx),".")[5];
//				System.out.println("项目编号str：" + okProjCode);
				
				
				//获得科目，在每组的第二行。用. split 之后 第3位是科目
				String okCostType = StringUtils.split(getCellValue2Trim(sheet, sIdx+1),".")[3];
//				System.out.println("科目str：" + okCostType);


				
				//获得没项目的明细中，每组本月合计的开始、结束索引
				List<Map<String,Integer>> byhjGroupIdxList = getSumIdxGroup(sheet, sIdx, eIdx);
				for(Map<String,Integer> idxMap : byhjGroupIdxList){
					
					Map<String,String> okMap = new HashMap<String,String>();
					okMap.put("proj_code", okProjCode);
					okMap.put("cost_type", okCostType);
					
					String colYMStr = getCellValue2Trim(sheet, idxMap.get("s"));
					String colMSumStr = getCellValue2Trim(sheet, idxMap.get("e"));
//					System.out.println("年月列str : " + colYMStr);
//					System.out.println("月合计列str : " + colMSumStr);
					//获得你年月，取colYMStr前8位后trim
					String[] ymVal = StringUtils.split(StringUtils.trimToEmpty(StringUtils.substring(colYMStr, 0, 8)),"-");
					String okYear = ymVal[0];
					String okMonth = ymVal[1];
					okMap.put("year", okYear);
					okMap.put("month", okMonth);
//					System.out.println("年:" + okYear + " - 月:" + okMonth);
					//处理月合计列，按空格split
					String[] sumVal = StringUtils.split(colMSumStr);
					String okJie = sumVal[1];
					String okDai = sumVal[2];
					String okYe = sumVal[4];
					okMap.put("jie", okJie);
					okMap.put("dai", okDai);
					okMap.put("ye", okYe);
//					System.out.println("借：" + okJie + " - 贷:" + okDai + " - 余额：" + okYe);
					okList.add(okMap);
				}
			}
			
//			System.out.println(okList);
			//write to xls
			toXls(okList);
			
			
		} catch (BiffException | IOException | WriteException e) {
			e.printStackTrace();
		}finally{
			workbook.close();
		}
	}

	private static void toXls(final List<Map<String, String>> okList) throws WriteException, IOException {
		WritableWorkbook workbook = null;
		try {
			workbook = Workbook.createWorkbook(new File("d://output"+System.currentTimeMillis()+".xls"));
			WritableSheet sheet = workbook.createSheet("sheet1", 0);
			//写表头
			sheet.addCell(new Label(1, 0, "项目编号"));
			sheet.addCell(new Label(2, 0, "成本类型"));
			sheet.addCell(new Label(3, 0, "年"));
			sheet.addCell(new Label(4, 0, "月"));
			sheet.addCell(new Label(5, 0, "借"));
			sheet.addCell(new Label(6, 0, "贷"));
			sheet.addCell(new Label(7, 0, "本月余额"));
			
			//写表格内容
			for(int i=0;i<okList.size();i++){
				Map<String,String> map = okList.get(i);
				sheet.addCell(new Label(1, i+1, map.get("proj_code")));
				sheet.addCell(new Label(2, i+1, map.get("cost_type")));
				sheet.addCell(new Label(3, i+1, map.get("year")));
				sheet.addCell(new Label(4, i+1, map.get("month")));
				sheet.addCell(new Label(5, i+1, map.get("jie")));
				sheet.addCell(new Label(6, i+1, map.get("dai")));
				sheet.addCell(new Label(7, i+1, map.get("ye")));
			}
			workbook.write();
		} finally{
			workbook.close();
		}
		
	}

	private static List<Map<String,Integer>> getSumIdxGroup(Sheet sheet, int sIdx, int eIdx) {
		//合计组明细的开始索引
		int sumGroupStartIdx = sIdx+4;
		
		//将包含“本月合计”的行打包,开始索引+4，去除头部信息。并将包含本月合计的索引保存起来
		List<Integer> byhjIdxList = new ArrayList<Integer>();
		for(int i=sumGroupStartIdx;i<=eIdx;i++){
			String byhjValStr = getCellValue2Trim(sheet, i);
//					System.out.println(byhjValStr);
			if(StringUtils.contains(byhjValStr, "本月合计")){
//						System.out.println("本月合计所在行：A"+i);
				byhjIdxList.add(i);
			}
		}
//				System.out.println(byhjIdxList);
		
		//处理包含本月合计的组开始索引和组结束索引
		List<Map<String,Integer>> byhjGroupIdxList = new ArrayList<Map<String,Integer>>();
		for(int i=0;i<byhjIdxList.size();i++){
			int e = byhjIdxList.get(i);
			Map<String,Integer> map = new HashMap<String,Integer>();
			map.put("s", sumGroupStartIdx);
			map.put("e", e);
			if(i+1 < byhjIdxList.size()){
				if(StringUtils.contains(getCellValue2Trim(sheet,e+1),"本年累计")){
					sumGroupStartIdx = e+2;
				}else{
					sumGroupStartIdx = e+1;
				}
			}
			byhjGroupIdxList.add(map);
		}
//		System.out.println(byhjGroupIdxList);
		return byhjGroupIdxList;
	}

	private static List<Map<String,Integer>> getGroupList(Sheet sheet, int totalRows) {
		//containKEMUList存放所有包含【科    目：】的行号
		List<Integer> containKEMUList = new ArrayList<Integer>();
		//获得文件中有多少组项目
		for(int i=1;i<=totalRows;i++){
			//获得col-A中的内容并去除空格
			String aVal = getCellValue2Trim(sheet,i);
			//将col-A中包含【科    目：】字样的行提取出来
			if(StringUtils.isNotBlank(aVal) && StringUtils.contains(aVal, "科    目：")){
//					System.out.println("A" + i + " - " + aVal);
				containKEMUList.add(i);
			}
		}
//			System.out.println(containKEMUList);
		//定义存放组数据的开始行s和结束行e的list
		List<Map<String,Integer>> groupList = new ArrayList<Map<String,Integer>>();
		for(int i=0;i<containKEMUList.size();i++){
			Map<String,Integer> groupMap = new HashMap<String,Integer>();
			groupMap.put("s", containKEMUList.get(i));
			//最后科目索引的结束等于总行数-2
			groupMap.put("e", i+1==containKEMUList.size() ? totalRows-2 : containKEMUList.get(i+1)-2);
			groupList.add(groupMap);
		}
//			System.out.println(groupList);
		return groupList;
	}
	
	//获得指定cell索引的内容并去除空格
	private static String getCellValue2Trim(Sheet sheet,int cellIdx){
		return StringUtils.trimToEmpty(sheet.getCell("A"+cellIdx).getContents());
	}

}
