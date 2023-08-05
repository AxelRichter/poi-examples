package arichter.examples.apache.poi.ss;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

/**
* ObjectCellValue provides methods to use {@link java.lang.Object} to set and get
* values of {@link org.apache.poi.ss.usermodel.Cell}  
*
* @author Axel Richter
* 
* @see org.apache.poi.ss.usermodel.Cell
* */
public class ObjectCellValue {
    
    private static final Logger LOG = LogManager.getLogger(ObjectCellValue.class);
    
    /**
    * Enum of usable classes to get date and/or date time values 
    * from {@link org.apache.poi.ss.usermodel.Cell} 
    */
    public enum DateTimeClass {
        /**
        * Use {@link java.util.Date} while getting date cell value
        */
        JAVA_UTIL_DATE, 
        /**
        * Use {@link java.time.LocalDateTime} while getting date cell value
        */        
        JAVA_TIME_LOCALDATETIME;
    }

    private DateTimeClass dateTimeClass = DateTimeClass.JAVA_UTIL_DATE;    
    
    /**
    * Constructor not used
    */
    public ObjectCellValue() {
    }

    
    /**
    * Setter method for used {@link DateTimeClass}
    * 
    * @param dateTimeClass the {@link DateTimeClass} used to get date and/or date time values
    * from {@link org.apache.poi.ss.usermodel.Cell}
    */
    public void setDateTimeClass(DateTimeClass dateTimeClass) {
        this.dateTimeClass = dateTimeClass;
    }   
    
     /**
    * Getter method for used {@link DateTimeClass}
    * 
    * @return {@link DateTimeClass} used to get date and/or date time values
    * from {@link org.apache.poi.ss.usermodel.Cell}
    */
    public DateTimeClass getDateTimeClass() {
        return this.dateTimeClass;
    }        
    
    /**
    * A method to get a {@link java.lang.Object} value from {@link org.apache.poi.ss.usermodel.Cell}
    * 
    * @param cell the {@link org.apache.poi.ss.usermodel.Cell} from which the value will be got
    * @param type the {@link org.apache.poi.ss.usermodel.CellType} of that cell
    * @return a {@link java.lang.Object} containing the cell value
    */
    public Object getCellValue(Cell cell, CellType type) {
        switch (type) {
            case STRING:
                return cell.getRichStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    switch (this.dateTimeClass) {
                        case JAVA_UTIL_DATE: return cell.getDateCellValue();
                        case JAVA_TIME_LOCALDATETIME: return cell.getLocalDateTimeCellValue();  
                    }
                    return cell.getDateCellValue();
                } else {
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                return getCellValue(cell, cell.getCachedFormulaResultType());
            case BLANK:
                return null;
            default:
                LOG.atWarn().log("This should not occur.", cell, type);
                return null;
        }
    }
    
    /**
    * A method to set a {@link java.lang.Object} value to a {@link org.apache.poi.ss.usermodel.Cell}
    * 
    * @param cell the {@link org.apache.poi.ss.usermodel.Cell} to which the value will be set
    * @param valueObject the {@link java.lang.Object} of bet set to that cell
    */
    public void setCellValue(Cell cell, Object valueObject) {
        if (valueObject instanceof String) {
            cell.setCellValue((String)valueObject);
        } else if (valueObject instanceof RichTextString) {
            cell.setCellValue((RichTextString)valueObject);
        } else if (valueObject instanceof Number) {
            cell.setCellValue(((Number)valueObject).doubleValue());
        } else if (valueObject instanceof Boolean) {
            cell.setCellValue((Boolean)valueObject);  
        } else if (valueObject instanceof java.util.Calendar) {
            cell.setCellValue((java.util.Calendar)valueObject);
            //use CellUtil to set the CellStyleProperty data format to date
            CellUtil.setCellStyleProperty(cell, CellUtil.DATA_FORMAT, 14);   
        } else if (valueObject instanceof java.util.Date) {
            cell.setCellValue((java.util.Date)valueObject);
            //use CellUtil to set the CellStyleProperty data format to date
            CellUtil.setCellStyleProperty(cell, CellUtil.DATA_FORMAT, 14);
        } else if (valueObject instanceof java.time.LocalDate) {
            cell.setCellValue((java.time.LocalDate)valueObject);
            //use CellUtil to set the CellStyleProperty data format to date
            CellUtil.setCellStyleProperty(cell, CellUtil.DATA_FORMAT, 14);
        } else if (valueObject instanceof java.time.LocalDateTime) {
            cell.setCellValue((java.time.LocalDateTime)valueObject);
            //use CellUtil to set the CellStyleProperty data format to date time
            CellUtil.setCellStyleProperty(cell, CellUtil.DATA_FORMAT, 22);
        } else if (valueObject == null) {
            String blank = null;
            cell.setCellValue(blank);            
        } else {
            cell.setCellValue(String.valueOf(valueObject)); 
        }
    }

}
