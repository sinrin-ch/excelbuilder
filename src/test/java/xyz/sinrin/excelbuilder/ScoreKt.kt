package xyz.sinrin.excelbuilder

import java.math.BigDecimal

data class ScoreKt(@CellConfig(0) var name: String = "", @CellConfig(1) var age: Int = 0, @CellConfig(2) var height: BigDecimal = 0.toBigDecimal()) {

}