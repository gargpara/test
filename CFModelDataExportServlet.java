package com.danone.aemaacs.core.servlets;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.jcr.RepositoryException;
import javax.jcr.Session;
import javax.servlet.Servlet;

import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.api.SlingHttpServletResponse;
import org.apache.sling.api.resource.Resource;
import org.apache.sling.api.resource.ResourceResolver;
import org.apache.sling.api.resource.ValueMap;
import org.apache.sling.api.servlets.HttpConstants;
import org.apache.sling.api.servlets.ServletResolverConstants;
import org.apache.sling.api.servlets.SlingAllMethodsServlet;
import org.apache.sling.jcr.resource.api.JcrResourceConstants;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.osgi.service.component.annotations.Component;
import org.osgi.service.component.annotations.Reference;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.adobe.cq.dam.cfm.ContentFragment;
import com.danone.aemaacs.core.utils.constants.Constants;
import com.day.cq.search.PredicateConverter;
import com.day.cq.search.PredicateGroup;
import com.day.cq.search.Query;
import com.day.cq.search.QueryBuilder;
import com.day.cq.search.eval.PathPredicateEvaluator;
import com.day.cq.search.eval.TypePredicateEvaluator;
import com.day.cq.search.result.Hit;
import com.day.cq.search.result.SearchResult;

@Component(service = Servlet.class, property = {
        ServletResolverConstants.SLING_SERVLET_NAME + "=Export CF Data as Excel",
        ServletResolverConstants.SLING_SERVLET_SELECTORS + "=export-cf",
        ServletResolverConstants.SLING_SERVLET_SELECTORS + "=export-cf-data",
        ServletResolverConstants.SLING_SERVLET_EXTENSIONS + "=xlsx",
        ServletResolverConstants.SLING_SERVLET_EXTENSIONS + "=json",
        ServletResolverConstants.SLING_SERVLET_METHODS + "=" + HttpConstants.METHOD_GET,
        ServletResolverConstants.SLING_SERVLET_RESOURCE_TYPES + "=cq/Page"
})
public class CFModelDataExportServlet extends SlingAllMethodsServlet {

    private static final long serialVersionUID = 1L;
    private static final Logger LOGGER = LoggerFactory.getLogger(CFModelDataExportServlet.class);

    @Reference
    private transient QueryBuilder queryBuilder;

    @Override
    protected void doGet(SlingHttpServletRequest req,
                         SlingHttpServletResponse resp)
                         throws IOException {

        String cfResource      = req.getParameter(Constants.CF_MODEL);
        String productsRoot    = req.getParameter(Constants.CF_ROOT_PATH);

        /* --- 1. Which variant did the author ask for? --- */
        boolean headerOnly = Arrays.asList(
                req.getRequestPathInfo().getSelectors()).contains("export-cf");

        /* --- 2. Build workbook ------------------------------------------------ */
        try (Workbook wb = new XSSFWorkbook()) {

            Sheet sheet   = wb.createSheet("CF Sample");
            List<String> headers = getHeaderAsList(req, cfResource);

            /* header row */
            Row headerRow = sheet.createRow(0);
            int i = 0;
            headerRow.createCell(i++).setCellValue("nodeName");
            for (String h : headers) headerRow.createCell(i++).setCellValue(h);

            /* data rows only if full export */
            if (!headerOnly) {
                try {
                    List<Resource> rows = getRelatedCFResource(
                                        cfResource, productsRoot, req.getResourceResolver(),
                                        req.getResourceResolver().adaptTo(Session.class));
                    addRowsToExcel(sheet, headers, rows);
                } catch (RepositoryException e) {
                    LOGGER.error("Failed to query related content fragments", e);
                }
            }

            /* --- 3. Stream ----------------------------------------------------- */
            String base = req.getResourceResolver()
                             .getResource(cfResource).getName();
            String file = base + (headerOnly ? "-sample" : "-data") + ".xlsx";

            resp.setContentType(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            resp.setHeader("Content-Disposition",
              "attachment; filename=\"" + file + '"');
            wb.write(resp.getOutputStream());
        }
    }

    /**
     * Builds the header list from the CF dialog resource.
     */
    private List<String> getHeaderAsList(SlingHttpServletRequest req, String cfResource) {

        List<String> header = new ArrayList<>();

        String dialogPath = cfResource + "/jcr:content/model/cq:dialog/content/items";
        Resource dialog   = req.getResourceResolver().getResource(dialogPath);

        if (dialog == null) {
            return header;
        }

        for (Resource child : dialog.getChildren()) {
            ValueMap vm    = child.getValueMap();
            String rt      = vm.get(JcrResourceConstants.SLING_RESOURCE_TYPE_PROPERTY, String.class);

            /*  skip tab placeholders (only when we can prove it is one) */
            if (rt != null && rt.endsWith("tabplaceholder")) {
                continue;
            }

            String fieldName = vm.get(Constants.NAME, String.class);
            if (fieldName != null && !fieldName.isBlank()) {
                header.add(fieldName);
            }
        }

        header.add("status");
        return header;
    }

    /**
     * Creates Excel rows (starting from row 1) based on the provided header and resource data.
     */
    private void addRowsToExcel(Sheet sheet, List<String> headerElements, List<Resource> rowElements) {
        int rowIndex = 1;
        for (Resource resource : rowElements) {
            Row row = sheet.createRow(rowIndex++);
            int cellIndex = 0;
            // First cell: the resource name.
            Cell cell = row.createCell(cellIndex++);
            cell.setCellValue(resource.getName());

            ContentFragment cf = resource.adaptTo(ContentFragment.class);
            String status = getPublishStatus(resource);
            // For each header (which includes the CF fields and the "status" column)
            for (String header : headerElements) {
                cell = row.createCell(cellIndex++);
                if ("status".equals(header)) {
                    cell.setCellValue(status);
                } else if (cf != null && cf.getElement(header) != null) {
                    // Handle multi-value fields.
                    if (cf.getElement(header).getValue().getValue() != null &&
                            cf.getElement(header).getValue().getDataType().isMultiValue()) {
                        String[] values = (String[]) cf.getElement(header).getValue().getValue();
                        if (values != null) {
                            StringBuilder cellContent = new StringBuilder();
                            for (String value : values) {
                                cellContent.append(value).append("|");
                            }
                            int length = cellContent.length() > 0 ? cellContent.length() - 1 : 0;
                            cell.setCellValue(cellContent.substring(0, length));
                        }
                    } else {
                        // Get the cell data.
                        String cellData = cf.getElement(header).getContent().toString().trim();
                        // Optionally: remove HTML tags so that rich HTML is not displayed.
                        cellData = cellData.replaceAll("<[^>]*>", "");
                        // Replace any newlines with a space.
                        cellData = cellData.replace("\n", " ").replace("\r", " ");
                        cell.setCellValue(cellData);
                    }
                } else {
                    cell.setCellValue("");
                }
            }
        }
    }

    /**
     * Retrieves the publish status of a resource based on its jcr:content replication property.
     */
    private String getPublishStatus(Resource asset) {
        if (asset == null) {
            return "Not published";
        }
        // Get the jcr:content node
        Resource content = asset.getChild("jcr:content");
        if (content == null) {
            return "Not published";
        }
        String action = content.getValueMap()
                             .get("cq:lastReplicationAction", String.class);
        if ("Activate".equalsIgnoreCase(action)) {
            return "Published";
        }
        if ("Deactivate".equalsIgnoreCase(action)) {
            return "Unpublished";
        }
        return "Not published";
    }

    /**
     * Retrieves related content fragments using the QueryBuilder API.
     */
    private List<Resource> getRelatedCFResource(String cfResource, String productsRootPath,
            ResourceResolver resourceResolver, Session session) throws RepositoryException {
        Map<String, String> predicateMap = createPredicateMap(productsRootPath, cfResource);
        PredicateGroup predicates = PredicateConverter.createPredicates(predicateMap);
        Query query = queryBuilder.createQuery(predicates, session);
        final SearchResult searchResult = query.getResult();
        final List<Resource> results = new ArrayList<>();
        LOGGER.info("getRelatedCFResource - {} results found", searchResult.getHits().size());
        for (final Hit hit : searchResult.getHits()) {
            results.add(resourceResolver.getResource(hit.getPath()));
        }
        return results;
    }

    /**
     * Creates the predicate map used for querying related content fragments.
     */
    private Map<String, String> createPredicateMap(String root, String cfModelTemplate) {
        Map<String, String> predicatesMap = new HashMap<>();
        predicatesMap.put(PathPredicateEvaluator.PATH, root);
        predicatesMap.put(TypePredicateEvaluator.TYPE, "dam:Asset");
        predicatesMap.put("p.limit", "-1");
        predicatesMap.put("group_1.property_1", "@jcr:content/contentFragment");
        predicatesMap.put("group_1.property.1_value", "true");
        predicatesMap.put("group_2.property_1", "@jcr:content/data/cq:model");
        predicatesMap.put("group_2.property-1.1_value", cfModelTemplate);
        return predicatesMap;
    }
}
