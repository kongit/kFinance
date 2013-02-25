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
			//����ļ���һ���ж�����
			int totalRows = sheet.getRows();
			System.out.println("��������" + totalRows);
			
			//���ݴ�����list
			List<Map<String,String>> okList = new ArrayList<Map<String,String>>();
			
			List<Map<String,Integer>> groupList = getGroupList(sheet, totalRows);
			//�õ�ÿ���������
			for(Map<String,Integer> groupMap : groupList){
				
				
				
				int sIdx = groupMap.get("s");
				int eIdx = groupMap.get("e");
				//�����Ŀ��ţ���ÿ��ĵ�һ�С���. split ֮�� ��5λ����Ŀ���
				String okProjCode = StringUtils.split(getCellValue2Trim(sheet, sIdx),".")[5];
//				System.out.println("��Ŀ���str��" + okProjCode);
				
				
				//��ÿ�Ŀ����ÿ��ĵڶ��С���. split ֮�� ��3λ�ǿ�Ŀ
				String okCostType = StringUtils.split(getCellValue2Trim(sheet, sIdx+1),".")[3];
//				System.out.println("��Ŀstr��" + okCostType);


				
				//���û��Ŀ����ϸ�У�ÿ�鱾�ºϼƵĿ�ʼ����������
				List<Map<String,Integer>> byhjGroupIdxList = getSumIdxGroup(sheet, sIdx, eIdx);
				for(Map<String,Integer> idxMap : byhjGroupIdxList){
					
					Map<String,String> okMap = new HashMap<String,String>();
					okMap.put("proj_code", okProjCode);
					okMap.put("cost_type", okCostType);
					
					String colYMStr = getCellValue2Trim(sheet, idxMap.get("s"));
					String colMSumStr = getCellValue2Trim(sheet, idxMap.get("e"));
//					System.out.println("������str : " + colYMStr);
//					System.out.println("�ºϼ���str : " + colMSumStr);
					//��������£�ȡcolYMStrǰ8λ��trim
					String[] ymVal = StringUtils.split(StringUtils.trimToEmpty(StringUtils.substring(colYMStr, 0, 8)),"-");
					String okYear = ymVal[0];
					String okMonth = ymVal[1];
					okMap.put("year", okYear);
					okMap.put("month", okMonth);
//					System.out.println("��:" + okYear + " - ��:" + okMonth);
					//�����ºϼ��У����ո�split
					String[] sumVal = StringUtils.split(colMSumStr);
					String okJie = sumVal[1];
					String okDai = sumVal[2];
					String okYe = sumVal[4];
					okMap.put("jie", okJie);
					okMap.put("dai", okDai);
					okMap.put("ye", okYe);
//					System.out.println("�裺" + okJie + " - ��:" + okDai + " - ��" + okYe);
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
			//д��ͷ
			sheet.addCell(new Label(1, 0, "��Ŀ���"));
			sheet.addCell(new Label(2, 0, "�ɱ�����"));
			sheet.addCell(new Label(3, 0, "��"));
			sheet.addCell(new Label(4, 0, "��"));
			sheet.addCell(new Label(5, 0, "��"));
			sheet.addCell(new Label(6, 0, "��"));
			sheet.addCell(new Label(7, 0, "�������"));
			
			//д�������
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
		//�ϼ�����ϸ�Ŀ�ʼ����
		int sumGroupStartIdx = sIdx+4;
		
		//�����������ºϼơ����д��,��ʼ����+4��ȥ��ͷ����Ϣ�������������ºϼƵ�������������
		List<Integer> byhjIdxList = new ArrayList<Integer>();
		for(int i=sumGroupStartIdx;i<=eIdx;i++){
			String byhjValStr = getCellValue2Trim(sheet, i);
//					System.out.println(byhjValStr);
			if(StringUtils.contains(byhjValStr, "���ºϼ�")){
//						System.out.println("���ºϼ������У�A"+i);
				byhjIdxList.add(i);
			}
		}
//				System.out.println(byhjIdxList);
		
		//����������ºϼƵ��鿪ʼ���������������
		List<Map<String,Integer>> byhjGroupIdxList = new ArrayList<Map<String,Integer>>();
		for(int i=0;i<byhjIdxList.size();i++){
			int e = byhjIdxList.get(i);
			Map<String,Integer> map = new HashMap<String,Integer>();
			map.put("s", sumGroupStartIdx);
			map.put("e", e);
			if(i+1 < byhjIdxList.size()){
				if(StringUtils.contains(getCellValue2Trim(sheet,e+1),"�����ۼ�")){
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
		//containKEMUList������а�������    Ŀ�������к�
		List<Integer> containKEMUList = new ArrayList<Integer>();
		//����ļ����ж�������Ŀ
		for(int i=1;i<=totalRows;i++){
			//���col-A�е����ݲ�ȥ���ո�
			String aVal = getCellValue2Trim(sheet,i);
			//��col-A�а�������    Ŀ��������������ȡ����
			if(StringUtils.isNotBlank(aVal) && StringUtils.contains(aVal, "��    Ŀ��")){
//					System.out.println("A" + i + " - " + aVal);
				containKEMUList.add(i);
			}
		}
//			System.out.println(containKEMUList);
		//�����������ݵĿ�ʼ��s�ͽ�����e��list
		List<Map<String,Integer>> groupList = new ArrayList<Map<String,Integer>>();
		for(int i=0;i<containKEMUList.size();i++){
			Map<String,Integer> groupMap = new HashMap<String,Integer>();
			groupMap.put("s", containKEMUList.get(i));
			//����Ŀ�����Ľ�������������-2
			groupMap.put("e", i+1==containKEMUList.size() ? totalRows-2 : containKEMUList.get(i+1)-2);
			groupList.add(groupMap);
		}
//			System.out.println(groupList);
		return groupList;
	}
	
	//���ָ��cell���������ݲ�ȥ���ո�
	private static String getCellValue2Trim(Sheet sheet,int cellIdx){
		return StringUtils.trimToEmpty(sheet.getCell("A"+cellIdx).getContents());
	}

}
