/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.calculatorexcel;

import java.io.IOException;
import javax.ws.rs.GET;
import javax.ws.rs.POST;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.QueryParam;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;

/**
 *
 * @author Nergal
 */
@Path("/rest")
public class Restinterface {
   
    @GET
    @Path("/echo")
    public String echo(@QueryParam("q") String original) {
        return original;
    }
    
    @GET
    @Path("/try")
    public String fileTrying() throws IOException {

        AnalyzeExcelFile af = new AnalyzeExcelFile();
        String rezult = af.analyze();
        return rezult;
    }
        
    /**
     * set excel file with request
     *
     * @param path - path
     */    
    @POST
    @Path("/file")
    @Produces({MediaType.TEXT_PLAIN})
//    @Consumes(MediaType.APPLICATION_JSON)    
    public Response setFile(String path){
        String result = "ok";
        return Response.ok(result).build();
    }
}
