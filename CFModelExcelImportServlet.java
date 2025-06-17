package com.danone.aemaacs.core.servlets;

import java.io.IOException;
import java.io.InputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.jcr.Node;
import javax.jcr.RepositoryException;
import javax.jcr.Session;
import javax.jcr.query.QueryManager;
import javax.servlet.Servlet;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.collections4.IteratorUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.sling.jcr.resource.api.JcrResourceConstants;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.api.SlingHttpServletResponse;
import org.apache.sling.api.request.RequestParameter;
import org.apache.sling.api.resource.LoginException;
import org.apache.sling.api.resource.PersistenceException;
import org.apache.sling.api.resource.Resource;
import org.apache.sling.api.resource.ResourceResolver;
import org.apache.sling.api.servlets.HttpConstants;
import org.apache.sling.api.servlets.ServletResolverConstants;
import org.osgi.service.component.annotations.Component;
import org.osgi.service.component.annotations.Reference;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.adobe.cq.dam.cfm.ContentElement;
import com.adobe.cq.dam.cfm.ContentFragment;
import com.adobe.cq.dam.cfm.ContentFragmentException;
import com.adobe.cq.dam.cfm.FragmentData;
import com.adobe.cq.dam.cfm.FragmentTemplate;
import com.danone.aemaacs.core.services.InstanceType;
import com.danone.aemaacs.core.utils.CommonUtils;
import com.danone.aemaacs.core.utils.ErrorUtils;
import com.danone.aemaacs.core.utils.constants.Constants;
import com.drew.lang.annotations.NotNull;
import com.google.gson.JsonArray;
import com.google.gson.JsonObject;

import lombok.extern.slf4j.Slf4j;

@Component(service = Servlet.class, property = {
        ServletResolverConstants.SLING_SERVLET_NAME + "=" + "Export CF model as Excel file",
        ServletResolverConstants.SLING_SERVLET_SELECTORS + "=" + "import-cf",
        ServletResolverConstants.SLING_SERVLET_EXTENSIONS + "=" + "json",
        ServletResolverConstants.SLING_SERVLET_RESOURCE_TYPES + "=" + "cq/Page",
        ServletResolverConstants.SLING_SERVLET_METHODS + "=" + HttpConstants.METHOD_POST 
})
@Slf4j
public class CFModelExcelImportServlet extends AbstractServlet {
    private static final String STATUS = "status";
    private static final Logger LOGGER = LoggerFactory.getLogger(CFModelExcelImportServlet.class);
    private static final long serialVersionUID = 1L;

    @Reference
    private transient InstanceType instanceType;

    @Override
    protected void doPost(@NotNull final SlingHttpServletRequest request, @NotNull final SlingHttpServletResponse response)
            throws IOException {
        if (!instanceType.isAuthorInstance()) {
            LOGGER.warn("Not an author instance, skipping import.");
            return;
        }

        boolean orverrideExisting = false;
        JsonArray jsonResponse = new JsonArray();

        try (ResourceResolver resolver = CommonUtils.retrieveResourceResolver()) {    
            // Retrieve parameters and log their values
            String productsRootPath = request.getParameter(Constants.CF_ROOT_PATH);
            String cfTemplate = request.getParameter(Constants.CF_MODEL);
            orverrideExisting = Boolean.parseBoolean(request.getParameter(Constants.OVERRIDE_EXISTING));
            LOGGER.info("Received parameters: productsRootPath={}, cfTemplate={}, overrideExisting={}",
                    productsRootPath, cfTemplate, orverrideExisting);
            if (StringUtils.isEmpty(productsRootPath) || StringUtils.isEmpty(cfTemplate)) {
                LOGGER.error("Missing required parameters: CF_ROOT_PATH or CF_MODEL");
                writeResponse(
                    response,
                    HttpServletResponse.SC_BAD_REQUEST,
                    "{\"error\":\"Missing required parameters\"}"
                );
                return;
            }
            // Retrieve the uploaded file parameter and log file details
            RequestParameter fileRequestParam = request.getRequestParameter(Constants.FILE);
            if (fileRequestParam == null || StringUtils.isBlank(fileRequestParam.getFileName())) {
                LOGGER.error("No file uploaded or file name is blank.");
                writeResponse(
                    response,
                    HttpServletResponse.SC_BAD_REQUEST,
                    "{\"error\":\"No file uploaded or file name is blank.\"}"
                );
                return;
            }
            LOGGER.info("File uploaded: {}", fileRequestParam.getFileName());
            InputStream is = fileRequestParam.getInputStream();
            if (is == null) {
                LOGGER.error("Input stream of uploaded file is null.");
                writeResponse(
                    response,
                    HttpServletResponse.SC_BAD_REQUEST,
                    "{\"error\":\"Input stream of uploaded file is null.\"}"
                );
                return;
            }

            // Convert the Excel workbook into a list of maps.
            List<String> modelHeaders = getModelHeaderList(resolver, cfTemplate);
            List<Map<String, Object>> excelRows = Collections.emptyList();
            try {
                excelRows = convertExcelToListOfMap(is, modelHeaders);
            } catch (IOException ioe) {
                LOGGER.error("Import failed while parsing Excel: {}", ioe.getMessage());
                writeResponse(
                    response,
                    HttpServletResponse.SC_BAD_REQUEST,
                    "{\"error\":\"" + ioe.getMessage() + "\"}"
                );
                return;
            }
            LOGGER.info("Parsed {} rows from Excel file.", excelRows.size());
            if (excelRows.isEmpty()) {
                LOGGER.error("No data rows found in the uploaded Excel file.");
                writeResponse(
                    response,
                    HttpServletResponse.SC_BAD_REQUEST,
                    "{\"error\":\"No data rows found in the uploaded Excel file.\"}"
                );
                return;
            }
            LOGGER.info("Headers from Excel: {}", excelRows.get(0).keySet());

            // Process the rows to create or update Content Fragments.
            jsonResponse = createCFsFromListOfHashMap(
                orverrideExisting, resolver,
                productsRootPath, cfTemplate, excelRows);

            try {
                resolver.commit();
            } catch (PersistenceException | RuntimeException ex) {
                LOGGER.error("Failed to commit changes", ex);
                writeResponse(response, HttpServletResponse.SC_INTERNAL_SERVER_ERROR,
                    "{\"error\":\"Repository commit failed: " + ex.getMessage() + "\"}");
                return;
            }
            LOGGER.info("Repository commit successful.");
        } catch (LoginException le) {
            LOGGER.error("Failed to get ResourceResolver.", le);
            writeResponse(
                response,
                HttpServletResponse.SC_BAD_REQUEST,
                "{\"error\":\"Failed to get ResourceResolver.\"}"
            );
            return;
        } catch (Exception e) {
            LOGGER.error("Unexpected import error.", e);
            writeResponse(
                response,
                HttpServletResponse.SC_BAD_REQUEST,
                "{\"error\":\"Unexpected import error.\"}"
            );
            return;
        }
        LOGGER.info("Import result:\n" + jsonResponse);
        writeResponse(
            response,
            HttpServletResponse.SC_OK,
            CommonUtils.simpleGSON.toJson(jsonResponse)
        );
    }

    /**
     * Converts the uploaded Excel (XLSX) file into a list of maps.
     * The first row is assumed to contain headers.
     */
    private List<Map<String, Object>> convertExcelToListOfMap(InputStream is,
                List<String> expectedHeaders) throws IOException {
        List<Map<String, Object>> rowsList = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(is)) {
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            // prepare a DataFormatter (and formula evaluator) up front:
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator eval = workbook.getCreationHelper().createFormulaEvaluator();

            // 1) Read headers:
            List<String> headerList = new ArrayList<>();
            if (rowIterator.hasNext()) {
                Row headerRow = rowIterator.next();
                for (Cell cell : headerRow) {
                    // draw the cell’s displayed text, no need to mutate its type
                    String header = fmt.formatCellValue(cell, eval);
                    headerList.add(header);
                }
                if (!headerList.equals(expectedHeaders)) {
                    throw new IOException("Columns mismatch. Expected: " + expectedHeaders);
                }
                LOGGER.info("Excel Headers: {}", headerList);
            } else {
                LOGGER.error("No header row found in Excel file.");
                throw new IOException("No header row found in Excel file.");
            }

            // 2) Read data rows:
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Map<String, Object> rowMap = new HashMap<>();
                for (int i = 0; i < headerList.size(); i++) {
                    Cell cell = row.getCell(i);
                    String cellValue = fmt.formatCellValue(cell, eval);
                    rowMap.put(headerList.get(i), cellValue);
                }
                rowsList.add(rowMap);
            }
        }
        return rowsList;
    }

    private JsonArray createCFsFromListOfHashMap(boolean orverrideExisting, ResourceResolver resourceResolver,
            String productsRootPath, String cfTemplate, List<Map<String, Object>> excelRows) {
        JsonArray jsonArrayResponse = new JsonArray();
        Resource templateOrModelRsc = resourceResolver.getResource(cfTemplate);
        Resource parentRsc = resourceResolver.getResource(productsRootPath);
        if (templateOrModelRsc != null) {
            Iterator<Resource> existingCFWithKey = Collections.emptyIterator();
            for (Map<String, Object> row : excelRows) {
                JsonObject jsonObject = new JsonObject();
                String keyValue = (row.get(Constants.KEY) != null) ? row.get(Constants.KEY).toString() : "undefined";
                LOGGER.info("Processing row with key: {}", keyValue);
                try {
                    jsonObject.addProperty("fragment_key", keyValue);
                    existingCFWithKey = findCFResources(keyValue, productsRootPath, cfTemplate, resourceResolver);
                    if (orverrideExisting && existingCFWithKey.hasNext()) {
                        LOGGER.info("Modifying existing CF for key: {}", keyValue);
                        modifyExistingCF(existingCFWithKey, row, jsonObject);
                    } else if (IteratorUtils.size(existingCFWithKey) < 1) {
                        LOGGER.info("Creating new CF for key: {}", keyValue);
                        canModifyCF(existingCFWithKey, keyValue);
                        createCF(parentRsc, templateOrModelRsc, row, jsonObject);
                    } else {
                        LOGGER.error("Multiple CF exist with key {}. Skipping creation.", keyValue);
                        List<String> cfPaths = new ArrayList<>();
                        while (existingCFWithKey.hasNext()) {
                            cfPaths.add(existingCFWithKey.next().getPath());
                        }
                        jsonObject.addProperty(STATUS, "Skipped");
                        jsonObject.addProperty("duplicates", cfPaths.toString());
                    }
                } catch (ContentFragmentException e) {
                    jsonObject.addProperty("error", e.getMessage());
                    ErrorUtils.loggerErrorDebug(LOGGER, "Unable to create Content Fragments",
                            getClass().getSimpleName(), e);
                }
                jsonArrayResponse.add(jsonObject);
            }
        } else {
            String errorMsg = "Content Fragment Model not found. Template path = " + cfTemplate;
            LOGGER.error(errorMsg);
            jsonArrayResponse.add(errorMsg);
        }
        return jsonArrayResponse;
    }

    private JsonObject createCF(Resource parentResource, Resource templateResource, Map<String, Object> model, JsonObject jsonObject)
            throws ContentFragmentException {
        String nodename = model.get("nodeName").toString();
        if (StringUtils.isEmpty(nodename)) {
            nodename = model.get("key").toString();
        }
        FragmentTemplate fragmentTemplate = templateResource.adaptTo(FragmentTemplate.class);
        if (fragmentTemplate != null) {
            ContentFragment cf = fragmentTemplate.createFragment(parentResource, nodename, nodename);
            JsonArray elements = writeFragmentData(model, cf);
            jsonObject.addProperty(STATUS, "Created");
            jsonObject.add("elements", elements);
            LOGGER.info("Created CF with name: {}", nodename);
        }
        return jsonObject;
    }

    private JsonObject modifyExistingCF(Iterator<Resource> existingCFWithKey, Map<String, Object> row, JsonObject jsonObject)
            throws ContentFragmentException {
        Resource rowResource = existingCFWithKey.next();
        ContentFragment fragment = rowResource.adaptTo(ContentFragment.class);
        if (fragment != null) {
            JsonArray elements = writeFragmentData(row, fragment);
            jsonObject.addProperty(STATUS, "Modified");
            jsonObject.add("elements", elements);
            LOGGER.info("Modified CF at path: {}", rowResource.getPath());
        } else {
            jsonObject.addProperty(STATUS, "Modified - issue while converting to fragment");
            LOGGER.error("Failed to adapt resource {} to ContentFragment.", rowResource.getPath());
        }
        return jsonObject;
    }

    private JsonArray writeFragmentData(Map<String, Object> row, ContentFragment cf)
            throws ContentFragmentException {
        JsonArray jsonArray = new JsonArray();
        Iterator<ContentElement> iteratorElements = cf.getElements();
        while (iteratorElements.hasNext()) {
            JsonObject jsonObject = new JsonObject();
            ContentElement elem = iteratorElements.next();
            LOGGER.info("Writing value for element '{}' of type '{}'", 
                    elem.getName(), elem.getValue().getDataType().getValueType().toString());
            FragmentData dataValue = elem.getValue();
            Object entry = row.get(elem.getName());
            if (entry != null) {
                LOGGER.info("Value from Excel for '{}': {}", elem.getName(), entry.toString());
                switch (elem.getValue().getDataType().getValueType().toString().toLowerCase()) {
                    case "string":
                        if (elem.getValue().getDataType().isMultiValue()) {
                            String[] value = processValue(entry.toString()).split("\\|");
                            dataValue.setValue(value);
                        } else {
                            dataValue.setValue(processValue(entry.toString()));
                        }
                        jsonObject.addProperty(elem.getName(), processValue(entry.toString()));
                        break;
                    case "long":
                        try {
                            int intValue = Integer.parseInt(entry.toString());
                            dataValue.setValue(intValue);
                            jsonObject.addProperty(elem.getName(), intValue);
                        } catch (NumberFormatException ex) {
                            jsonObject.addProperty(elem.getName(), "Issue while parsing integer");
                        }
                        break;
                    case "double":
                        try {
                            double doubleValue = Double.parseDouble(entry.toString());
                            dataValue.setValue(doubleValue);
                            jsonObject.addProperty(elem.getName(), doubleValue);
                        } catch (NumberFormatException ex) {
                            jsonObject.addProperty(elem.getName(), "Issue while parsing double");
                        }
                        break;
                    case "calendar":
                        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
                        try {
                            Date date = formatter.parse(entry.toString());
                            Calendar calendar = Calendar.getInstance();
                            calendar.setTime(date);
                            dataValue.setValue(calendar);
                            jsonObject.addProperty(elem.getName(), entry.toString());
                        } catch (ParseException e) {
                            jsonObject.addProperty(elem.getName(), "Issue while converting to date");
                        }
                        break;
                    default:
                        jsonObject.addProperty(elem.getName(), StringUtils.EMPTY);
                }
            } else {
                jsonObject.addProperty(elem.getName(), StringUtils.EMPTY);
            }
            elem.setValue(dataValue);
            jsonArray.add(jsonObject);
        }
        return jsonArray;
    }

    private boolean canModifyCF(Iterator<Resource> resources, String key) {
        if (resources != null && IteratorUtils.size(resources) > 1) {
            LOGGER.error("Multiple CF found with key: {}", key);
            while (resources.hasNext()) {
                LOGGER.error("Existing CF path: {}", resources.next().getPath());
            }
            return false;
        }
        return true;
    }

    private Iterator<Resource> findCFResources(String key, String path, String model, ResourceResolver resourceResolver) {
        String query = "SELECT * FROM [dam:Asset] WHERE [jcr:content/contentFragment]=true AND [jcr:content/data/cq:model]='"
                + model + "' AND [jcr:content/data/master/key]='" + key + "' AND ISCHILDNODE('" + path + "')";
        List<Resource> allResources = new ArrayList<>();
        try {
            Session session = resourceResolver.adaptTo(Session.class);
            if (session != null) {
                QueryManager queryManager = session.getWorkspace().getQueryManager();
                javax.jcr.query.Query q = queryManager.createQuery(query, javax.jcr.query.Query.JCR_SQL2);
                @SuppressWarnings("unchecked")
                Iterator<Node> nodes = q.execute().getNodes();
                while (nodes.hasNext()) {
                    Node node = nodes.next();
                    Resource resource = resourceResolver.getResource(node.getPath());
                    if (resource != null) {
                        allResources.add(resource);
                    }
                }
            }
            LOGGER.info("Found {} CF resources for key: {}", allResources.size(), key);
            return allResources.iterator();
        } catch (RepositoryException e) {
            LOGGER.error("Problem retrieving resources for key: {}", key, e);
        }
        return Collections.emptyIterator();
    }

    private String processValue(String value) {
        if (value.contains("<nl>")) {
            return value.replaceAll("<nl>", "<br>");
        }
        return value;
    }

    private List<String> getModelHeaderList(ResourceResolver resolver, String cfTemplate) {
        List<String> headers = new ArrayList<>();
        headers.add("nodeName");
        String dialogPath = cfTemplate + "/jcr:content/model/cq:dialog/content/items";
        Resource dialogRes = resolver.getResource(dialogPath);
        if (dialogRes != null) {
            Iterator<Resource> children = dialogRes.listChildren();
            while (children.hasNext()) {
                Resource child = children.next();
                String type = child.getValueMap()
                                   .get(JcrResourceConstants.SLING_RESOURCE_TYPE_PROPERTY, "");
                if (type.endsWith("tabplaceholder")) {
                    continue;
                }
                headers.add(child.getValueMap()
                                 .get(Constants.NAME, ""));
            }
            // match your import’s “status” column
            headers.add("status");
        }
        return headers;
    }
    
}
