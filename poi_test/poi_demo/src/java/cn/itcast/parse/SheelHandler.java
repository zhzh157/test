package cn.itcast.parse;

import cn.itcast.domain.ContractProductVo;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.text.ParseException;
import java.text.SimpleDateFormat;

public class SheelHandler implements XSSFSheetXMLHandler.SheetContentsHandler  {
   private ContractProductVo vo;
    public void startRow(int i) {
        if (i>=2){
            vo=new ContractProductVo();
        }
    }

    public void endRow(int i) {
        System.out.println(vo);
    }

    public void cell(String cellName, String cellValue, XSSFComment xssfComment) {
        String name=cellName.substring(0,1);
        if (vo!=null){
            /*switch (name) {
                case "B" :{
                    vo.setCustomName(cellValue);
                    break;
                }
                case "C" :{
                    vo.setContractNo(cellValue);
                    break;
                }
                case "D" :{
                    vo.setProductNo(cellValue);
                    break;
                }
                case "E" :{
                    vo.setCnumber(Integer.parseInt(cellValue));
                    break;
                }
                case "F" :{
                    vo.setFactoryName(cellValue);
                    break;
                }
                case "G" :{

                    try {
                        vo.setDeliveryPeriod(new SimpleDateFormat("yyyy-MM-dd").parse(cellValue) );
                    } catch (ParseException e) {
                        e.printStackTrace();
                    }

                    break;
                }
                case "H" :{

                    try {
                        vo.setShipTime(new SimpleDateFormat("yyyy-MM-dd").parse(cellValue) );
                    } catch (ParseException e) {
                        e.printStackTrace();
                    }

                    break;
                }
                case "I" :{
                    vo.setTradeTerms(cellValue);
                    break;
                }
                default:{
                    break;
                }
            }*/
        }
    }
}
