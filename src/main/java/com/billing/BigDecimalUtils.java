package com.billing;

import org.threeten.extra.Temporals;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.text.DecimalFormatSymbols;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.temporal.TemporalAdjusters;
import java.util.ArrayList;
import java.util.List;

import lombok.Generated;

import static java.math.BigDecimal.ZERO;
import static java.math.RoundingMode.HALF_UP;
import static java.time.DayOfWeek.SATURDAY;
import static java.time.DayOfWeek.SUNDAY;
import static java.time.LocalDate.now;

public abstract class BigDecimalUtils {

    /**
     * Always use compareTo for comparing BigDecimals as equals
     * returns true only if the two compared objects have the same scale
     */
    public static boolean isZero(BigDecimal amount) {

        if (amount == null) return true;
        return amount.compareTo(ZERO) == 0;
    }

    public static boolean isNotZero(BigDecimal amount) {

        return !isZero(amount);
    }

    public static boolean areEqual(BigDecimal first, BigDecimal second) {
        return first.compareTo(second) == 0;
    }

    /**
     * Formats a BigDecimal with a scale of exactly 2
     *
     * @param bigDecimal to be formatted
     * @return formatted bigDecimal as string
     */
    public static String format(BigDecimal bigDecimal, int decimalDigits) {

        if (isZero(bigDecimal)) return "0.00";

        bigDecimal = bigDecimal.setScale(decimalDigits, HALF_UP);

        var decimalFormat = new DecimalFormat();

        decimalFormat.setMaximumFractionDigits(decimalDigits);
        decimalFormat.setMinimumFractionDigits(decimalDigits);
        decimalFormat.setGroupingUsed(true);

        var decimalFormatSymbols = DecimalFormatSymbols.getInstance();
        decimalFormatSymbols.setDecimalSeparator(',');
        decimalFormatSymbols.setGroupingSeparator('.');
        decimalFormat.setDecimalFormatSymbols(decimalFormatSymbols);

        return decimalFormat.format(bigDecimal);
    }

    public static void main(String... args) {
        var firstOfMonth = now().with(TemporalAdjusters.firstDayOfMonth());
        var nextWorkingDay = firstOfMonth.with(Temporals.nextWorkingDayOrSame());
        var workingDays = new ArrayList<LocalDate>();
        while(nextWorkingDay.getMonthValue() == now().getMonthValue()){
            workingDays.add(nextWorkingDay);
            nextWorkingDay = nextWorkingDay.with(Temporals.nextWorkingDay());
        }
        workingDays.forEach(System.out::println);

    }
}
