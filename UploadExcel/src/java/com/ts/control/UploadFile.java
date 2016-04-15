/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
//File upload
//         |
//         v
//http://stackoverflow.com/questions/22985809/upload-read-an-excel-file-in-a-jsp-using-poi
package com.ts.control;

import java.io.*;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.commons.fileupload.FileItem;
import org.apache.commons.fileupload.FileItemFactory;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.commons.fileupload.servlet.ServletFileUpload;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author AlejandroA
 */
public class UploadFile extends HttpServlet {
    String saveFile="C:/lol"; 
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        response.setContentType("text/html;charset=UTF-8");
        PrintWriter out = response.getWriter();
        String campo1="";
        String campo2="";
        String campo3="";
        String campo4="";
        String campo5="";
        String campo6="";
        String campo7="";
        int c=1;
            try {
                boolean ismultipart=ServletFileUpload.isMultipartContent(request);
                if(!ismultipart){
                }else{
                    FileItemFactory factory = new DiskFileItemFactory();
                    ServletFileUpload upload = new ServletFileUpload(factory);
                    List items = null;
                    try{
                    items = upload.parseRequest(request);
                    }catch(Exception e){
                    }
                    Iterator itr = items.iterator();
                    while(itr.hasNext()){
                    FileItem item = (FileItem)itr.next();
                    if(item.isFormField()){
                    }else{
                    String itemname = item.getName();
                    if((itemname==null || itemname.equals(""))){
                    continue;
                    }
                    String filename = FilenameUtils.getName(itemname);
                    System.out.println("ruta---------------"+saveFile+filename);
                    File f = checkExist(filename);
                    item.write(f);
                    ////////
                    List cellDataList = new ArrayList();
                    try{
                    FileInputStream fileInputStream = new FileInputStream(f);
                    XSSFWorkbook workBook = new XSSFWorkbook(fileInputStream);
                    XSSFSheet hssfSheet = workBook.getSheetAt(0);
                    Iterator rowIterator = hssfSheet.rowIterator();
                    while(rowIterator.hasNext()){
                    XSSFRow hssfRow = (XSSFRow) rowIterator.next();
                    Iterator iterator = hssfRow.cellIterator();
                    List cellTempList = new ArrayList();
                    while(iterator.hasNext()){
                    XSSFCell hssfCell = (XSSFCell) iterator.next();
                    cellTempList.add(hssfCell);
                    }
                    cellDataList.add(cellTempList);
                    }
                    }catch (Exception e)
                    {
                    e.printStackTrace();
                    }
                    out.println("<table border='3' class='table table-bordered table-striped table-hover'>");
                          for (int i = 1; i < cellDataList.size(); i++){
                                List cellTempList = (List) cellDataList.get(i);
                                out.println("<tr>");
                                    for (int j = 0; j < cellTempList.size(); j++){
                                    XSSFCell hssfCell = (XSSFCell) cellTempList.get(j);
                                    String dato = hssfCell.toString();
                                    out.print("<td>"+dato+"</td>");
                                    switch(c){
                                        case 1:
                                            campo1=dato;
                                            c++;
                                            break;
                                        case 2:
                                            campo2=dato;
                                            c++;
                                            break;
                                        case 3:
                                            campo3=dato;
                                            c++;
                                            break; 
                                        case 4:
                                            campo4=dato;
                                            c++;
                                            break;
                                        case 5:
                                            campo5=dato;
                                            c++;
                                            break;
                                        case 6:
                                            campo6=dato;
                                            c++;
                                            break;
                                       case 7:
                                            campo7=dato;
                                            c=1;
                                            break;     
                                    }
                                    }
                                    //model.Consulta.addRegistros(campo1,campo2,campo3,campo4,campo5,campo6,campo7); 
                                }
                                 out.println("</table><br>");
                    /////////
                    }
                    }
                }
                }catch(Exception e){
                }
               finally {
                out.close();
                }
        }
    
        private File checkExist(String fileName) {
        File f = new File(saveFile+"/"+fileName);

        if(f.exists()){
        StringBuffer sb = new StringBuffer(fileName);
        sb.insert(sb.lastIndexOf("."),"-"+new Date().getTime());
        f = new File(saveFile+"/"+sb.toString());
        }
        return f;
        }

    // <editor-fold defaultstate="collapsed" desc="HttpServlet methods. Click on the + sign on the left to edit the code.">
    /**
     * Handles the HTTP <code>GET</code> method.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        processRequest(request, response);
    }

    /**
     * Handles the HTTP <code>POST</code> method.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        processRequest(request, response);
    }

    /**
     * Returns a short description of the servlet.
     *
     * @return a String containing servlet description
     */
    @Override
    public String getServletInfo() {
        return "Short description";
    }// </editor-fold>

}
