/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package isi.tecnologias.ods;

import javax.xml.xpath.*;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import javax.xml.parsers.DocumentBuilderFactory;
import org.odftoolkit.simple.table.Cell;
import org.odftoolkit.simple.SpreadsheetDocument;
import org.odftoolkit.simple.table.Table;

/**
 *
 * @author vitor
 */

public class ISITecnologiasODS {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        String UI="CIIC", XMLFolder="CVs-CIIC", xmlFile="cv-all.xml", ODSFolder="ODS-CIIC", odsFile="CIIC-Indicadores.ods";
        SpreadsheetDocument outputOds;
        try { // Esta linha apenas retira informação de Java Logging da consola
                 java.util.logging.LogManager.getLogManager().readConfiguration(new java.io.ByteArrayInputStream("org.odftoolkit.level=WARNING".getBytes(java.nio.charset.StandardCharsets.UTF_8)));
            outputOds = SpreadsheetDocument.loadDocument("./Indicadores-template.ods");      
            Table template = outputOds.getSheetByIndex(0);            
            outputOds.insertSheet(template, 0);                            
            Table folhaNova = outputOds.getSheetByIndex(0);                        
            folhaNova.setTableName("ano 1");                          
            Document CvAllXMLDoc = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(XMLFolder + "/" + xmlFile);
            NodeList InvestigadoresNodeList = (NodeList) XPathFactory.newInstance().newXPath().compile("//DADOS-GERAIS/@NOME-COMPLETO").evaluate(CvAllXMLDoc, XPathConstants.NODESET);
            double contarTrabalhosFactor=0;
            double contarTrabalhosSemFactor=0;
            double factorMedio=0;
            double conferenciaIndexada=0;
            String nomeDoInvestigador;
            double numerosCapitulos=0;
            double numerosLivros=0;
            double numerosRevista=0;
            
            for (int i=0; i < InvestigadoresNodeList.getLength(); i++) {
                nomeDoInvestigador = InvestigadoresNodeList.item(i).getFirstChild().getNodeValue();         
                contarTrabalhosFactor = (double) XPathFactory.newInstance().newXPath().compile("count(//CURRICULO-VITAE/DADOS-GERAIS[@NOME-COMPLETO='"+nomeDoInvestigador+"']/../PRODUCAO-BIBLIOGRAFICA/ARTIGOS-PUBLICADOS/ARTIGO-PUBLICADO/DETALHAMENTO-DO-ARTIGO[@ARBITRAGEM-CIENTIFICA='SIM' and @FACTOR-DE-IMPACTO-JCR]/..)").evaluate(CvAllXMLDoc, XPathConstants.NUMBER);
                contarTrabalhosSemFactor = (double) XPathFactory.newInstance().newXPath().compile("count(//CURRICULO-VITAE/DADOS-GERAIS[@NOME-COMPLETO='"+nomeDoInvestigador+"']/../PRODUCAO-BIBLIOGRAFICA/ARTIGOS-PUBLICADOS/ARTIGO-PUBLICADO/DETALHAMENTO-DO-ARTIGO[@ARBITRAGEM-CIENTIFICA='SIM']/..)").evaluate(CvAllXMLDoc, XPathConstants.NUMBER);
                factorMedio = (double) XPathFactory.newInstance().newXPath().compile("sum(//CURRICULO-VITAE/DADOS-GERAIS[@NOME-COMPLETO='"+nomeDoInvestigador+"']/../PRODUCAO-BIBLIOGRAFICA/ARTIGOS-PUBLICADOS/ARTIGO-PUBLICADO/DETALHAMENTO-DO-ARTIGO/@FACTOR-DE-IMPACTO-JCR)").evaluate(CvAllXMLDoc, XPathConstants.NUMBER);
                conferenciaIndexada = (double) XPathFactory.newInstance().newXPath().compile("count(//CURRICULO-VITAE/DADOS-GERAIS[@NOME-COMPLETO='"+nomeDoInvestigador+"']/../PRODUCAO-TECNICA/DEMAIS-TIPOS-DE-PRODUCAO-TECNICA/APRESENTACAO-DE-TRABALHO/DADOS-BASICOS-DA-APRESENTACAO-DE-TRABALHO[@NATUREZA='CONFERENCIA']/..)").evaluate(CvAllXMLDoc, XPathConstants.NUMBER);
                numerosCapitulos = (double) XPathFactory.newInstance().newXPath().compile("count(//CURRICULO-VITAE/DADOS-GERAIS[@NOME-COMPLETO='"+nomeDoInvestigador+"']/../PRODUCAO-BIBLIOGRAFICA/LIVROS-E-CAPITULOS/CAPITULOS-DE-LIVROS-PUBLICADOS)").evaluate(CvAllXMLDoc, XPathConstants.NUMBER);
                numerosLivros = (double) XPathFactory.newInstance().newXPath().compile("count(//CURRICULO-VITAE/DADOS-GERAIS[@NOME-COMPLETO='"+nomeDoInvestigador+"']/../PRODUCAO-BIBLIOGRAFICA/LIVROS-E-CAPITULOS/LIVROS-PUBLICADOS-OU-ORGANIZADOS)").evaluate(CvAllXMLDoc, XPathConstants.NUMBER);
                numerosRevista = (double) XPathFactory.newInstance().newXPath().compile("count(//CURRICULO-VITAE/DADOS-GERAIS[@NOME-COMPLETO='"+nomeDoInvestigador+"']/../DADOS-COMPLEMENTARES/PARTICIPACAO-EDITORIAL-EM-REVISTAS/PARTICIPACAO-EDITORIAL-EM-REVISTA)").evaluate(CvAllXMLDoc, XPathConstants.NUMBER);
                
                factorMedio= factorMedio/contarTrabalhosFactor;
                if (contarTrabalhosFactor==0)
                    factorMedio=0;
                Cell cell1 = folhaNova.getCellByPosition(1,i+5);
                Cell cell2 = folhaNova.getCellByPosition(10,i+5);
                Cell cell3 = folhaNova.getCellByPosition(11,i+5);
                Cell cell4 = folhaNova.getCellByPosition(9,i+5);
                Cell cell5 = folhaNova.getCellByPosition(12,i+5);
                Cell cell6 = folhaNova.getCellByPosition(13,i+5);
                Cell cell7 = folhaNova.getCellByPosition(14,i+5);
                Cell cell8 = folhaNova.getCellByPosition(15,i+5);
                
                
                cell1.setStringValue(nomeDoInvestigador);
                cell2.setDoubleValue(contarTrabalhosFactor);
                cell3.setDoubleValue(contarTrabalhosSemFactor);
                cell4.setDoubleValue(factorMedio);
                cell5.setDoubleValue(conferenciaIndexada);
                cell6.setDoubleValue(numerosCapitulos);
                cell7.setDoubleValue(numerosLivros);
                cell8.setDoubleValue(numerosRevista);
            }            
            System.out.println("A gravar folha " + ODSFolder+"/"+odsFile);
            outputOds.save(ODSFolder+"/"+odsFile);
        } catch (Exception e) {
            System.err.println("ERRO: " + e.getMessage() + " " + e.toString());
        }                                
    }
}
