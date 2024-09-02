package com.excel.receipt.model;

import lombok.Data;

@Data
public class Item {
    String name;
    int quantity;
    double price;

    public double getTotalPrice() {
        return quantity * price;
    }
}
