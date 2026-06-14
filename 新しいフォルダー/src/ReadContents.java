package com.example.sql_generator;

public class ReadContents {
	/** 項目論理名 */
	String colNmJp;
	/** 項目物理名(Snake) */
	String colNmSnk;
	/** 項目物理名(Camel) */
	String colNmCml;
	/** 項目型 */
	String colType;

	public ReadContents(String colNmJp, String colNmSnk, String colNmCml, String colType) {
		this.colNmJp = colNmJp;
		this.colNmSnk = colNmSnk;
		this.colNmCml = colNmCml;
		this.colType = colType;
	}

	public String getColNmJp() {
		return colNmJp;
	}

	public void setColNmJp(String colNmJp) {
		this.colNmJp = colNmJp;
	}

	public String getColNmSnk() {
		return colNmSnk;
	}

	public void setColNmSnk(String colNmSnk) {
		this.colNmSnk = colNmSnk;
	}

	public String getColNmCml() {
		return colNmCml;
	}

	public void setColNmCml(String colNmCml) {
		this.colNmCml = colNmCml;
	}

	public String getColType() {
		return colType;
	}

	public void setColType(String colType) {
		this.colType = colType;
	}
}
