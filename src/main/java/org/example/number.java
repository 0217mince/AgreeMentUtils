package org.example;

import java.math.BigDecimal;

public class number {
    public static void main(String[] args) {
        BigDecimal bigDecimal = new BigDecimal(0.00);
        System.out.println(bigDecimal.stripTrailingZeros().compareTo(BigDecimal.ZERO) == 0);
    }
}
