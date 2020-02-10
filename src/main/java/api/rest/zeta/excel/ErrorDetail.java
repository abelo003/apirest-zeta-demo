/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package api.rest.zeta.excel;

import java.io.Serializable;

/**
 * Clase que contiene informacion del error encontrado
 * @author Abelo
 */
public class ErrorDetail implements Serializable{
    
    private String buc;
    private String pan;

    /**
     * Contructor para objeto error
     * @param buc buc del cliente
     * @param pan pan del cliente
     */
    public ErrorDetail(String buc, String pan) {
        this.buc = buc;
        this.pan = pan;
    }

    public String getBuc() {
        return buc;
    }

    public void setBuc(String buc) {
        this.buc = buc;
    }

    public String getPan() {
        return pan;
    }

    public void setPan(String pan) {
        this.pan = pan;
    }
    
}
