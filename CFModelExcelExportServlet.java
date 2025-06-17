package com.danone.aemaacs.core.servlets;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.UnsupportedEncodingException;
import java.util.Iterator;

import javax.servlet.Servlet;

import org.apache.sling.api.SlingHttpServletRequest;
import org.apache.sling.api.SlingHttpServletResponse;
import org.apache.sling.api.resource.Resource;
import org.apache.sling.api.resource.ResourceResolver;
import org.apache.sling.api.servlets.HttpConstants;
import org.apache.sling.api.servlets.ServletResolverConstants;
import org.apache.sling.jcr.resource.api.JcrResourceConstants;
import org.osgi.service.component.annotations.Component;

import com.danone.aemaacs.core.utils.constants.Constants;

import lombok.extern.slf4j.Slf4j;

@Component(service = Servlet.class, property = {
        ServletResolverConstants.SLING_SERVLET_NAME + "=" + "Export CF model as CSV file",
        ServletResolverConstants.SLING_SERVLET_SELECTORS + "=" + "export-cf",
        ServletResolverConstants.SLING_SERVLET_EXTENSIONS + "=" + "json",
        ServletResolverConstants.SLING_SERVLET_RESOURCE_TYPES + "=" + "cq/Page",
        ServletResolverConstants.SLING_SERVLET_METHODS + "=" + HttpConstants.METHOD_GET
})
@Slf4j
public class CFModelExcelExportServlet extends AbstractServlet {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	
	@Override
    protected void doGet(final SlingHttpServletRequest req, final SlingHttpServletResponse resp) throws UnsupportedEncodingException, IOException {
		BufferedWriter writer = new BufferedWriter(new OutputStreamWriter(resp.getOutputStream(), Constants.UTF_8));
		StringBuilder cfResource = new StringBuilder(req.getParameter(Constants.CF_MODEL).toString());
		cfResource.append("/jcr:content/model/cq:dialog/content/items");
		ResourceResolver resolver = req.getResourceResolver();
		Resource dialogResource = resolver.getResource(cfResource.toString());
		if(dialogResource!=null) {
			Iterator<Resource> dialogchild = dialogResource.listChildren();
			StringBuilder content = new StringBuilder("nodeName;");
			while(dialogchild.hasNext()) {
				Resource child = dialogchild.next();
				if(child.getValueMap().get(JcrResourceConstants.SLING_RESOURCE_TYPE_PROPERTY).toString().endsWith("tabplaceholder")) {
					continue;
				} else {
					content.append(child.getValueMap().get(Constants.NAME).toString()+";");
				}
			}
			String value = content.toString().substring(0, content.toString().length()-1);
			writer.append(value);
			writer.newLine();
		}
		
		StringBuilder fileName =  new StringBuilder("sample-");
		Resource modelResouce = resolver.getResource(req.getParameter(Constants.CF_MODEL).toString());
		if(modelResouce !=null) {
			fileName = fileName.append(modelResouce.getName());
		}
		
		resp.setHeader("Content-Type", Constants.TEXT_CSV);
	    resp.setHeader("Content-Disposition", "attachment;filename=\""+fileName+".csv\"");
	    writer.flush();
		
	}
}
