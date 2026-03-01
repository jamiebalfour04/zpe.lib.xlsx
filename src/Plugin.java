/*
 * zpe.lib.xlsx
 *
 * Apache POI based XLSX support for ZPE/YASS.
 *
 * Global functions:
 *   - xlsx_new() => ZPEXLSXWorkbook
 *   - xlsx_open(string path) => ZPEXLSXWorkbook | false
 *
 * Objects:
 *   - ZPEXLSXWorkbook (workbook)
 *   - ZPEXLSXSheet (sheet)
 *
 * Permissions (suggested):
 *   - In-memory creation: 0
 *   - File open/save: 3
 */

import jamiebalfour.HelperFunctions;
import jamiebalfour.generic.JBBinarySearchTree;
import jamiebalfour.zpe.core.*;
import jamiebalfour.zpe.interfaces.*;
import jamiebalfour.zpe.types.ZPEBoolean;
import jamiebalfour.zpe.types.ZPENumber;
import jamiebalfour.zpe.types.ZPEString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

public class Plugin implements ZPELibrary {

  @Override
  public Map<String, ZPECustomFunction> getFunctions() {
    HashMap<String, ZPECustomFunction> arr = new HashMap<>();
    arr.put("xlsx_new", new XLSXNew());
    arr.put("xlsx_open", new XLSXOpen());
    return arr;
  }

  /**
   * Returning your object classes here helps ZPE expose them cleanly (e.g. for manual generation / reflection).
   */
  @Override
  public Map<String, Class<? extends ZPEStructure>> getObjects() {
    HashMap<String, Class<? extends ZPEStructure>> obj = new HashMap<>();
    obj.put("ZPEXLSXWorkbook", ZPEXLSXWorkbookObject.class);
    obj.put("ZPEXLSXSheet", ZPEXLSXSheetObject.class);
    return obj;
  }

  @Override
  public boolean supportsWindows() {
    return true;
  }

  @Override
  public boolean supportsMacOs() {
    return true;
  }

  @Override
  public boolean supportsLinux() {
    return true;
  }

  @Override
  public String getName() {
    return "libXLSX";
  }

  @Override
  public String getVersionInfo() {
    return "1.0";
  }

  // =============================================================================
  // Global function: xlsx_new()
  // =============================================================================
  public static final class XLSXNew implements ZPECustomFunction {

    @Override
    public String getManualEntry() {
      return "Creates a new XLSX workbook in memory.";
    }

    @Override
    public String getManualHeader() {
      return "xlsx_new ([])";
    }

    @Override
    public int getMinimumParameters() {
      return 0;
    }

    @Override
    public String[] getParameterNames() {
      return new String[]{};
    }

    @Override
    public ZPEType MainMethod(HashMap<String, Object> params, ZPERuntimeEnvironment runtime, ZPEFunction fn) {
      try {
        ZPEXLSXWorkbookObject wb = new ZPEXLSXWorkbookObject(runtime, fn);
        wb.newFile();
        return wb;
      } catch (Exception e) {
        return new ZPEBoolean(false);
      }
    }

    @Override
    public int getRequiredPermissionLevel() {
      return 0;
    }

    @Override
    public byte[] getReturnTypes() {
      // Workbook object (structure)
      return new byte[]{YASSByteCodes.OBJECT};
    }
  }

  // =============================================================================
  // Global function: xlsx_open(path)
  // =============================================================================
  public static final class XLSXOpen implements ZPECustomFunction {

    @Override
    public String getManualEntry() {
      return "Opens an XLSX workbook from disk.";
    }

    @Override
    public String getManualHeader() {
      return "xlsx_open ([{string} path])";
    }

    @Override
    public int getMinimumParameters() {
      return 1;
    }

    @Override
    public String[] getParameterNames() {
      return new String[]{"path"};
    }

    @Override
    public ZPEType MainMethod(HashMap<String, Object> params, ZPERuntimeEnvironment runtime, ZPEFunction fn) {
      try {
        String path = (params.get("path") == null) ? "" : params.get("path").toString();
        if (path.isEmpty()) return new ZPEBoolean(false);

        ZPEXLSXWorkbookObject wb = new ZPEXLSXWorkbookObject(runtime, fn);
        return wb.open(path) ? wb : new ZPEBoolean(false);

      } catch (Exception e) {
        return new ZPEBoolean(false);
      }
    }

    @Override
    public int getRequiredPermissionLevel() {
      return 3;
    }

    @Override
    public byte[] getReturnTypes() {
      // workbook | false
      return new byte[]{YASSByteCodes.OBJECT, YASSByteCodes.BOOLEAN_TYPE};
    }
  }

  // =============================================================================
  // ZPEXLSXWorkbookObject
  // =============================================================================
  public static final class ZPEXLSXWorkbookObject extends ZPEStructure {

    private static final long serialVersionUID = 3341840321048823111L;

    private transient XSSFWorkbook workbook;

    public ZPEXLSXWorkbookObject(ZPERuntimeEnvironment z, ZPEPropertyWrapper parent) {
      super(z, parent, "ZPEXLSXWorkbook");

      addNativeMethod("new_file", new new_file_Command());
      addNativeMethod("open", new open_Command());
      addNativeMethod("save", new save_Command());
      addNativeMethod("close", new close_Command());

      addNativeMethod("add_sheet", new add_sheet_Command());
      addNativeMethod("get_sheet", new get_sheet_Command());
      addNativeMethod("get_sheet_count", new get_sheet_count_Command());
    }

    void newFile() {
      closeQuietly();
      workbook = new XSSFWorkbook();
      workbook.createSheet("Sheet1");
    }

    boolean open(String path) {
      closeQuietly();
      try (FileInputStream fis = new FileInputStream(path)) {
        workbook = new XSSFWorkbook(fis);
        return true;
      } catch (Exception e) {
        workbook = null;
        return false;
      }
    }

    boolean save(String path) {
      if (workbook == null) return false;
      try (FileOutputStream fos = new FileOutputStream(path)) {
        workbook.write(fos);
        fos.flush();
        return true;
      } catch (Exception e) {
        return false;
      }
    }

    boolean close() {
      return closeQuietly();
    }

    private boolean closeQuietly() {
      try {
        if (workbook != null) {
          workbook.close();
        }
        workbook = null;
        return true;
      } catch (Exception e) {
        workbook = null;
        return false;
      }
    }

    XSSFWorkbook getWorkbook() {
      return workbook;
    }

    // ----------------------------
    // Native methods
    // ----------------------------

    static final class new_file_Command implements ZPEObjectNativeMethod {
      @Override
      public String[] getParameterNames() {
        return new String[]{};
      }

      @Override
      public String[] getParameterTypes() {
        return new String[]{};
      }

      @Override
      public ZPEType MainMethod(JBBinarySearchTree<String, ZPEType> parameters, ZPEObject parent) {
        ((ZPEXLSXWorkbookObject) parent).newFile();
        return parent;
      }

      @Override
      public int getRequiredPermissionLevel() {
        return 0;
      }

      @Override
      public String getName() {
        return "new_file";
      }

      @Override
      public byte[] returnTypes() {
        return new byte[]{YASSByteCodes.OBJECT};
      }
    }

    static final class open_Command implements ZPEObjectNativeMethod {
      @Override
      public String[] getParameterNames() {
        return new String[]{"path"};
      }

      @Override
      public String[] getParameterTypes() {
        return new String[]{"string"};
      }

      @Override
      public ZPEType MainMethod(JBBinarySearchTree<String, ZPEType> parameters, ZPEObject parent) {
        try {
          String path = parameters.get("path").toString();
          return new ZPEBoolean(((ZPEXLSXWorkbookObject) parent).open(path));
        } catch (Exception e) {
          return new ZPEBoolean(false);
        }
      }

      @Override
      public int getRequiredPermissionLevel() {
        return 3;
      }

      @Override
      public String getName() {
        return "open";
      }

      @Override
      public byte[] returnTypes() {
        return new byte[]{YASSByteCodes.BOOLEAN_TYPE};
      }
    }

    static final class save_Command implements ZPEObjectNativeMethod {
      @Override
      public String[] getParameterNames() {
        return new String[]{"path"};
      }

      @Override
      public String[] getParameterTypes() {
        return new String[]{"string"};
      }

      @Override
      public ZPEType MainMethod(JBBinarySearchTree<String, ZPEType> parameters, ZPEObject parent) {
        try {
          String path = parameters.get("path").toString();
          return new ZPEBoolean(((ZPEXLSXWorkbookObject) parent).save(path));
        } catch (Exception e) {
          return new ZPEBoolean(false);
        }
      }

      @Override
      public int getRequiredPermissionLevel() {
        return 3;
      }

      @Override
      public String getName() {
        return "save";
      }

      @Override
      public byte[] returnTypes() {
        return new byte[]{YASSByteCodes.BOOLEAN_TYPE};
      }
    }

    static final class close_Command implements ZPEObjectNativeMethod {
      @Override
      public String[] getParameterNames() {
        return new String[]{};
      }

      @Override
      public String[] getParameterTypes() {
        return new String[]{};
      }

      @Override
      public ZPEType MainMethod(JBBinarySearchTree<String, ZPEType> parameters, ZPEObject parent) {
        return new ZPEBoolean(((ZPEXLSXWorkbookObject) parent).close());
      }

      @Override
      public int getRequiredPermissionLevel() {
        return 0;
      }

      @Override
      public String getName() {
        return "close";
      }

      @Override
      public byte[] returnTypes() {
        return new byte[]{YASSByteCodes.BOOLEAN_TYPE};
      }
    }

    static final class add_sheet_Command implements ZPEObjectNativeMethod {
      @Override
      public String[] getParameterNames() {
        return new String[]{"name"};
      }

      @Override
      public String[] getParameterTypes() {
        return new String[]{"string"};
      }

      @Override
      public ZPEType MainMethod(JBBinarySearchTree<String, ZPEType> parameters, ZPEObject parent) {
        try {
          ZPEXLSXWorkbookObject wb = (ZPEXLSXWorkbookObject) parent;
          if (wb.getWorkbook() == null) return new ZPEBoolean(false);

          String name = parameters.get("name").toString();
          if (name.trim().isEmpty()) name = "Sheet" + (wb.getWorkbook().getNumberOfSheets() + 1);

          XSSFSheet sheet = wb.getWorkbook().createSheet(name);
          return new ZPEXLSXSheetObject(wb.getRuntime(), wb, wb, sheet);

        } catch (Exception e) {
          return new ZPEBoolean(false);
        }
      }

      @Override
      public int getRequiredPermissionLevel() {
        return 0;
      }

      @Override
      public String getName() {
        return "add_sheet";
      }

      @Override
      public byte[] returnTypes() {
        return new byte[]{YASSByteCodes.OBJECT, YASSByteCodes.BOOLEAN_TYPE};
      }
    }

    static final class get_sheet_Command implements ZPEObjectNativeMethod {
      @Override
      public String[] getParameterNames() {
        return new String[]{"name_or_index"};
      }

      @Override
      public String[] getParameterTypes() {
        return new String[]{"mixed"};
      }

      @Override
      public ZPEType MainMethod(JBBinarySearchTree<String, ZPEType> parameters, ZPEObject parent) {
        try {
          ZPEXLSXWorkbookObject wb = (ZPEXLSXWorkbookObject) parent;
          if (wb.getWorkbook() == null) return new ZPEBoolean(false);

          ZPEType v = parameters.get("name_or_index");
          if (v == null) return new ZPEBoolean(false);

          XSSFSheet sheet = null;

          String s = v.toString();
          Integer idx = null;
          try {
            idx = HelperFunctions.stringToInteger(s);
          } catch (Exception ignored) {
          }

          if (idx != null) {
            if (idx < 0 || idx >= wb.getWorkbook().getNumberOfSheets()) return new ZPEBoolean(false);
            sheet = wb.getWorkbook().getSheetAt(idx);
          } else {
            sheet = wb.getWorkbook().getSheet(s);
          }

          if (sheet == null) return new ZPEBoolean(false);
          return new ZPEXLSXSheetObject(wb.getRuntime(), wb, wb, sheet);

        } catch (Exception e) {
          return new ZPEBoolean(false);
        }
      }

      @Override
      public int getRequiredPermissionLevel() {
        return 0;
      }

      @Override
      public String getName() {
        return "get_sheet";
      }

      @Override
      public byte[] returnTypes() {
        return new byte[]{YASSByteCodes.OBJECT, YASSByteCodes.BOOLEAN_TYPE};
      }
    }

    static final class get_sheet_count_Command implements ZPEObjectNativeMethod {
      @Override
      public String[] getParameterNames() {
        return new String[]{};
      }

      @Override
      public String[] getParameterTypes() {
        return new String[]{};
      }

      @Override
      public ZPEType MainMethod(JBBinarySearchTree<String, ZPEType> parameters, ZPEObject parent) {
        try {
          ZPEXLSXWorkbookObject wb = (ZPEXLSXWorkbookObject) parent;
          if (wb.getWorkbook() == null) return new ZPEBoolean(false);
          return new ZPENumber(wb.getWorkbook().getNumberOfSheets());
        } catch (Exception e) {
          return new ZPEBoolean(false);
        }
      }

      @Override
      public int getRequiredPermissionLevel() {
        return 0;
      }

      @Override
      public String getName() {
        return "get_sheet_count";
      }

      @Override
      public byte[] returnTypes() {
        return new byte[]{YASSByteCodes.NUMBER_TYPE, YASSByteCodes.BOOLEAN_TYPE};
      }
    }
  }

  // =============================================================================
  // ZPEXLSXSheetObject
  // =============================================================================
  public static final class ZPEXLSXSheetObject extends ZPEStructure {

    private static final long serialVersionUID = 7412849723950412345L;

    private final ZPEXLSXWorkbookObject workbookObj;
    private final transient XSSFSheet sheet;

    public ZPEXLSXSheetObject(ZPERuntimeEnvironment z, ZPEPropertyWrapper parent, ZPEXLSXWorkbookObject workbookObj, XSSFSheet sheet) {
      super(z, parent, "ZPEXLSXSheet");
      this.workbookObj = workbookObj;
      this.sheet = sheet;

      addNativeMethod("set_cell", new set_cell_Command());
      addNativeMethod("get_cell", new get_cell_Command());
      addNativeMethod("get_last_row", new get_last_row_Command());
      addNativeMethod("get_name", new get_name_Command());
    }

    private static int asInt(ZPEType t) {
      return HelperFunctions.stringToInteger(t.toString());
    }

    private Row ensureRow(int rowIndex) {
      Row r = sheet.getRow(rowIndex);
      if (r == null) r = sheet.createRow(rowIndex);
      return r;
    }

    private Cell ensureCell(Row r, int colIndex) {
      Cell c = r.getCell(colIndex);
      if (c == null) c = r.createCell(colIndex);
      return c;
    }

    // ----------------------------
    // Native methods
    // ----------------------------

    final class set_cell_Command implements ZPEObjectNativeMethod {
      @Override
      public String[] getParameterNames() {
        return new String[]{"row", "col", "value"};
      }

      @Override
      public String[] getParameterTypes() {
        return new String[]{"number", "number", "mixed"};
      }

      @Override
      public ZPEType MainMethod(JBBinarySearchTree<String, ZPEType> parameters, ZPEObject parent) {
        try {
          int row = asInt(parameters.get("row"));
          int col = asInt(parameters.get("col"));
          ZPEType value = parameters.get("value");

          if (row < 0 || col < 0) return new ZPEBoolean(false);

          Row r = ensureRow(row);
          Cell c = ensureCell(r, col);

          String vs = (value == null) ? "" : value.toString();

          // Boolean?
          if ("true".equalsIgnoreCase(vs) || "false".equalsIgnoreCase(vs)) {
            c.setCellValue(Boolean.parseBoolean(vs));
            return new ZPEBoolean(true);
          }

          // Number?
          try {
            double d = Double.parseDouble(vs.trim());
            c.setCellValue(d);
            return new ZPEBoolean(true);
          } catch (Exception ignored) {
          }

          // Default string
          c.setCellValue(vs);
          return new ZPEBoolean(true);

        } catch (Exception e) {
          return new ZPEBoolean(false);
        }
      }

      @Override
      public int getRequiredPermissionLevel() {
        return 0;
      }

      @Override
      public String getName() {
        return "set_cell";
      }

      @Override
      public byte[] returnTypes() {
        return new byte[]{YASSByteCodes.BOOLEAN_TYPE};
      }
    }

    final class get_cell_Command implements ZPEObjectNativeMethod {
      @Override
      public String[] getParameterNames() {
        return new String[]{"row", "col"};
      }

      @Override
      public String[] getParameterTypes() {
        return new String[]{"number", "number"};
      }

      @Override
      public ZPEType MainMethod(JBBinarySearchTree<String, ZPEType> parameters, ZPEObject parent) {
        try {
          int row = asInt(parameters.get("row"));
          int col = asInt(parameters.get("col"));

          if (row < 0 || col < 0) return new ZPEBoolean(false);

          Row r = sheet.getRow(row);
          if (r == null) return new ZPEString("");

          Cell c = r.getCell(col);
          if (c == null) return new ZPEString("");

          CellType ct = c.getCellType();
          if (ct == CellType.FORMULA) {
            // Keep your old behaviour: return the formula string
            return new ZPEString(c.getCellFormula());
          }

          switch (ct) {
            case STRING:
              return new ZPEString(c.getStringCellValue());
            case BOOLEAN:
              return new ZPEBoolean(c.getBooleanCellValue());
            case NUMERIC:
              return new ZPENumber(c.getNumericCellValue());
            case BLANK:
            default:
              return new ZPEString("");
          }

        } catch (Exception e) {
          return new ZPEBoolean(false);
        }
      }

      @Override
      public int getRequiredPermissionLevel() {
        return 0;
      }

      @Override
      public String getName() {
        return "get_cell";
      }

      @Override
      public byte[] returnTypes() {
        return new byte[]{YASSByteCodes.STRING_TYPE, YASSByteCodes.NUMBER_TYPE, YASSByteCodes.BOOLEAN_TYPE};
      }
    }

    final class get_last_row_Command implements ZPEObjectNativeMethod {
      @Override
      public String[] getParameterNames() {
        return new String[]{};
      }

      @Override
      public String[] getParameterTypes() {
        return new String[]{};
      }

      @Override
      public ZPEType MainMethod(JBBinarySearchTree<String, ZPEType> parameters, ZPEObject parent) {
        try {
          return new ZPENumber(sheet.getLastRowNum());
        } catch (Exception e) {
          return new ZPEBoolean(false);
        }
      }

      @Override
      public int getRequiredPermissionLevel() {
        return 0;
      }

      @Override
      public String getName() {
        return "get_last_row";
      }

      @Override
      public byte[] returnTypes() {
        return new byte[]{YASSByteCodes.NUMBER_TYPE, YASSByteCodes.BOOLEAN_TYPE};
      }
    }

    final class get_name_Command implements ZPEObjectNativeMethod {
      @Override
      public String[] getParameterNames() {
        return new String[]{};
      }

      @Override
      public String[] getParameterTypes() {
        return new String[]{};
      }

      @Override
      public ZPEType MainMethod(JBBinarySearchTree<String, ZPEType> parameters, ZPEObject parent) {
        try {
          return new ZPEString(sheet.getSheetName());
        } catch (Exception e) {
          return new ZPEBoolean(false);
        }
      }

      @Override
      public int getRequiredPermissionLevel() {
        return 0;
      }

      @Override
      public String getName() {
        return "get_name";
      }

      @Override
      public byte[] returnTypes() {
        return new byte[]{YASSByteCodes.STRING_TYPE, YASSByteCodes.BOOLEAN_TYPE};
      }
    }
  }
}