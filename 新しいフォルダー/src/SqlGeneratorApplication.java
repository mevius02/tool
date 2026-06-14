package com.example.sql_generator;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class SqlGeneratorApplication {

	public static final String DIR_TABLE = "generateTable/";
	public static final String DIR_ENTITY = "generateEntity/";
	public static final String DIR_MAPPER = "generateMapper/";
	public static final String DIR_XML = "generateXml/";
	public static final String ENTITY = "entity";
	public static final String MAPPER = "mapper";
	public static final String XML = "xml";
	public static final String CHARACTER = "character";
	public static final String INTEGER = "integer";
	public static final String TIMESTAMP = "timestamp";
	public static final String DATE = "date";
	public static final String BYTEA = "bytea";
	public static final String ARRAY = "array";

	public static void main(String[] args) {
		SpringApplication.run(SqlGeneratorApplication.class, args);

		// フォルダ[generateTable]内の全ファイル名を取得
		File generateTableList = new File(DIR_TABLE);
		String[] generateTableArr = generateTableList.list();

		// ファイル数繰り返し
		for (int i = 0; i < generateTableArr.length; i++) {
			System.out.println("★彡 【" + generateTableArr[i] + "】のファイルを作成中... ★彡");

			// テーブル物理名取得(ファイル名から拡張子切り抜きしただけ)
			String tblNm = generateTableArr[i].substring(0, generateTableArr[i].lastIndexOf("."));
			// テーブル論理名
			String tblNmJp = null;
			// PK(Snake)
			String pkNmSnk = null;

			// ファイル内容取得
			// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
			int row = 1;
			List<ReadContents> readContentsList = new ArrayList<>();
			try (BufferedReader br = new BufferedReader(
					new FileReader(DIR_TABLE + new File(generateTableArr[i]), Charset.forName("UTF-8")))) {
				String line;
				while ((line = br.readLine()) != null) {
					// タブ区切り
					String[] values = line.split("\t", -1);
					// １行目：[テーブル論理名]
					if (1 == row) {
						tblNmJp = values[0];
					}
					// ２行目：[PK]
					if (2 == row) {
						pkNmSnk = values[0];
					}
					// ３行目以降：[項目論理名], [項目物理名(Sneke)], [項目物理名(Camel)], [項目型]
					// 項目型の分析、セット
					if (2 < row) {
						String type = null;
						if (values[2].startsWith(CHARACTER)) {
							// 配列型の場合
							if (values[2].endsWith("[]")) {
								type = ARRAY;
							} else {
								type = CHARACTER;
							}
						} else if (values[2].startsWith(INTEGER)) {
							type = INTEGER;
						} else if (values[2].startsWith(TIMESTAMP) || values[2].startsWith(DATE)) {
							type = DATE;
						} else if (values[2].startsWith(BYTEA)) {
							type = BYTEA;
						}
						readContentsList.add(
								new ReadContents(values[0], values[1], convertCamelCase(values[1], false), type));
					}
					row++;
				}
			} catch (IOException e) {
				System.out.println("ファイル読み込みエラー");
				System.out.println(e);
				lastShow(false, 10000);
				System.exit(0);
			}

			// ファイル作成
			// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
			createFile(ENTITY, tblNm, tblNmJp, pkNmSnk, readContentsList);
			createFile(MAPPER, tblNm, tblNmJp, pkNmSnk, readContentsList);
			createFile(XML, tblNm, tblNmJp, pkNmSnk, readContentsList);
		}
		lastShow(true, 3000);
		System.exit(0);
	}

	/**
	 * ファイル作成
	 * 
	 * @param createType
	 * @param tblNm
	 * @param tblNmJp
	 * @param pkNmSnk
	 * @param readContentsList
	 */
	private static void createFile(String createType, String tblNm, String tblNmJp, String pkNmSnk,
			List<ReadContents> readContentsList) {
		// 作成パス
		String currentPath = null;
		// 作成ファイル名
		String createFileNm = null;
		// 書き込み内容
		String writeContents = null;
		// ファイル作成
		try {
			switch (createType) {
				case ENTITY:
					currentPath = DIR_ENTITY;
					createFileNm = convertCamelCase(tblNm, true) + "Entity.java";
					writeContents = createContentsEntity(tblNm, tblNmJp, readContentsList);
					break;
				case MAPPER:
					currentPath = DIR_MAPPER;
					createFileNm = convertCamelCase(tblNm, true) + "Mapper.java";
					writeContents = createContentsMapper(tblNm, tblNmJp, readContentsList);
					break;
				case XML:
					currentPath = DIR_XML;
					createFileNm = convertCamelCase(tblNm, true) + "Mapper.xml";
					writeContents = createContentsXml(tblNm, tblNmJp, pkNmSnk, readContentsList);
					break;
			}
			// フォルダ存在しない場合作成
			File createPath = new File(currentPath + convertCamelCase(tblNm, false));
			if (!createPath.exists()) {
				createPath.mkdirs();
			}
			// 書き込み
			FileWriter fw = new FileWriter(createPath + "/" + createFileNm, Charset.forName("UTF-8"));
			fw.write(writeContents);
			fw.close();
		} catch (IOException e) {
			System.out.println("ファイル書き込みエラー");
			System.out.println(e);
			lastShow(true, 10000);
			System.exit(0);
		}
	}

	/**
	 * スネークケースからキャメルケース変換
	 * 
	 * @param snakeCaseStr
	 * @param initialUpperFlg 頭文字大文字フラグ[TRUE=大文字/FALSE=小文字]
	 * @return
	 */
	private static String convertCamelCase(String snakeCaseStr, boolean initialUpperFlg) {
		String[] parts = snakeCaseStr.split("_");
		StringBuilder sb = new StringBuilder();
		for (String part : parts) {
			sb.append(part.substring(0, 1).toUpperCase() + part.substring(1));
		}
		String result = sb.toString();
		if (!initialUpperFlg) {
			result = result.substring(0, 1).toLowerCase() + result.substring(1);
		}
		return result;
	}

	/**
	 * 内容作成(Entity)
	 * 
	 * @param tblNm
	 * @param tblNmJp
	 * @param readContentsList
	 * @return
	 */
	private static String createContentsEntity(String tblNm, String tblNmJp, List<ReadContents> readContentsList) {
		StringBuilder sb = new StringBuilder();

		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		// PACKAGE
		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		sb.append("package jp.progmat.ut.common.domain.entity." + convertCamelCase(tblNm, false) + ";\n");
		sb.append("\n");
		sb.append("import java.util.Date;\n");
		sb.append("\n");
		sb.append("import jakarta.persistence.Entity;\n");
		sb.append("import jakarta.persistence.Id;\n");
		sb.append("import lombok.Getter;\n");
		sb.append("import lombok.Setter;\n");
		sb.append("\n");
		sb.append("/** " + tblNmJp + " */\n");
		sb.append("@Entity\n");
		sb.append("@Getter\n");
		sb.append("@Setter\n");
		sb.append("public class " + convertCamelCase(tblNm, true) + "Entity {\n");
		sb.append("\n");
		boolean first = true;
		for (ReadContents data : readContentsList) {
			sb.append("\t/** " + data.getColNmJp() + " */\n");
			if (first) {
				sb.append("\t@Id\n");
				first = false;
			}
			if (data.getColType().equals(CHARACTER)) {
				sb.append("\tprivate String " + data.getColNmCml() + ";\n");
			} else if (data.getColType().equals(INTEGER)) {
				sb.append("\tprivate Integer " + data.getColNmCml() + ";\n");
			} else if (data.getColType().equals(TIMESTAMP) || data.getColType().equals(DATE)) {
				sb.append("\tprivate Date " + data.getColNmCml() + ";\n");
			} else if (data.getColType().equals(BYTEA)) {
				sb.append("\tprivate byte[] " + data.getColNmCml() + ";\n");
			} else if (data.getColType().equals(ARRAY)) {
				sb.append("\tprivate Object " + data.getColNmCml() + ";\n");
			}
		}
		sb.append("\n");
		sb.append("\t/** コンストラクター */\n");
		sb.append("\tpublic " + convertCamelCase(tblNm, true) + "Entity() {}\n");
		sb.append("\n");
		sb.append("\t/**\n");
		sb.append("\t * コンストラクター\n");
		int blankNum = 0;
		for (ReadContents data : readContentsList) {
			if (blankNum < data.getColNmCml().length()) {
				blankNum = data.getColNmCml().length();
			}
		}
		String blank = new String(new char[blankNum + 1]).replace("\0", " ");
		for (ReadContents data : readContentsList) {
			sb.append("\t * @param " + data.getColNmCml()
					+ blank.substring(data.getColNmCml().length()) + data.getColNmJp() + "\n");
		}
		sb.append("\t */\n");
		sb.append("\tpublic " + convertCamelCase(tblNm, true) + "Entity(");
		for (ReadContents data : readContentsList) {
			if (data.getColType().equals(CHARACTER)) {
				sb.append("String " + data.getColNmCml() + ", ");
			} else if (data.getColType().equals(INTEGER)) {
				sb.append("Integer " + data.getColNmCml() + ", ");
			} else if (data.getColType().equals(TIMESTAMP) || data.getColType().equals(DATE)) {
				sb.append("Date " + data.getColNmCml() + ", ");
			} else if (data.getColType().equals(BYTEA)) {
				sb.append("byte[] " + data.getColNmCml() + ", ");
			} else if (data.getColType().equals(ARRAY)) {
				sb.append("Object " + data.getColNmCml() + ", ");
			}
		}
		sb = new StringBuilder(sb.substring(0, sb.length() - 2));
		sb.append(") {\n");
		for (ReadContents data : readContentsList) {
			sb.append("\t\tthis." + data.getColNmCml() + " = " + data.getColNmCml() + ";\n");
		}
		sb.append("\t}\n");
		sb.append("}\n");
		return sb.toString();
	}

	/**
	 * 内容作成(Mapper)
	 * 
	 * @param tblNm
	 * @param tblNmJp
	 * @param readContentsList
	 * @return
	 */
	private static String createContentsMapper(String tblNm, String tblNmJp, List<ReadContents> readContentsList) {
		StringBuilder sb = new StringBuilder();

		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		// PACKAGE
		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		sb.append("package jp.progmat.ut.common.infra.mapper." + convertCamelCase(tblNm, false) + ";\n");
		sb.append("\n");
		sb.append("import java.util.List;\n");
		sb.append("\n");
		sb.append("import org.apache.ibatis.annotations.Mapper;\n");
		sb.append("import org.apache.ibatis.annotations.Param;\n");
		sb.append("\n");
		sb.append("import jp.progmat.ut.common.domain.entity." + convertCamelCase(tblNm, false) + "."
				+ convertCamelCase(tblNm, true) + "Entity;\n");
		sb.append("\n");
		sb.append("/** " + tblNmJp + " */\n");
		sb.append("@Mapper\n");
		sb.append("public interface " + convertCamelCase(tblNm, true) + "Mapper {\n");
		sb.append("\n");
		sb.append("\t/**\n");
		sb.append("\t * [検索]" + tblNmJp + "\n");
		sb.append("\t * @param distinct DISTINCT [TRUE=DISTINCT付与する/FALSE=DISTINCT付与しない]\n");
		sb.append("\t * @param where    WHERE    イコール比較する項目のみ値が入っているEntityクラス\n");
		sb.append("\t * @param orderBy  ORDER BY ORDER BY句の内容\n");
		sb.append("\t * @return 検索結果リスト\n");
		sb.append("\t */\n");
		sb.append("\t List<" + convertCamelCase(tblNm, true)
				+ "Entity> select(@Param(\"distinct\") boolean distinct, @Param(\"where\") "
				+ convertCamelCase(tblNm, true) + "Entity where, @Param(\"orderBy\") String orderBy);\n");
		sb.append("\n");
		sb.append("\t/**\n");
		sb.append("\t * [検索]" + tblNmJp + "\n");
		sb.append("\t * @param distinct DISTINCT [TRUE=DISTINCT付与する/FALSE=DISTINCT付与しない]\n");
		sb.append("\t * @param where    WHERE    イコール比較する項目のみ値が入っているEntityクラス\n");
		sb.append("\t * @param orderBy  ORDER BY ORDER BY句の内容\n");
		sb.append("\t * @return 検索結果件数\n");
		sb.append("\t */\n");
		sb.append(
				"\t long count(@Param(\"distinct\") boolean distinct, @Param(\"where\") "
						+ convertCamelCase(tblNm, true) + "Entity where, @Param(\"orderBy\") String orderBy);\n");
		sb.append("\n");
		sb.append("\t/**\n");
		sb.append("\t * [追加]" + tblNmJp + "\n");
		sb.append("\t * @param record レコード\n");
		sb.append("\t * @return 追加件数\n");
		sb.append("\t */\n");
		sb.append("\t int insert(@Param(\"record\") " + convertCamelCase(tblNm, true) + "Entity record);\n");
		sb.append("\n");
		sb.append("\t/**\n");
		sb.append("\t * [追加]" + tblNmJp + "(選択項目)\n");
		sb.append("\t * <pre>説明：引数[record]内の各変数値がnull以外の項目のみ対象となる</pre>\n");
		sb.append("\t * <pre>つまり、nullで追加、更新ができない</pre>\n");
		sb.append("\t * @param record レコード\n");
		sb.append("\t * @return 追加件数\n");
		sb.append("\t */\n");
		sb.append("\t int insertSelective(@Param(\"record\") " + convertCamelCase(tblNm, true) + "Entity record);\n");
		sb.append("\n");
		sb.append("\t/**\n");
		sb.append("\t * [更新]" + tblNmJp + "\n");
		sb.append("\t * @param record レコード\n");
		sb.append("\t * @param where  WHERE イコール比較する項目のみ値が入っているEntityクラス\n");
		sb.append("\t * @return 更新件数\n");
		sb.append("\t */\n");
		sb.append("\t int update(@Param(\"record\") " + convertCamelCase(tblNm, true)
				+ "Entity record, @Param(\"where\") " + convertCamelCase(tblNm, true) + "Entity where);\n");
		sb.append("\n");
		sb.append("\t/**\n");
		sb.append("\t * [更新]" + tblNmJp + "(選択項目)\n");
		sb.append("\t * <pre>説明：引数[record]内の各変数値がnull以外の項目のみ対象となる</pre>\n");
		sb.append("\t * <pre>つまり、nullで追加、更新ができない</pre>\n");
		sb.append("\t * @param record レコード\n");
		sb.append("\t * @param where  WHERE イコール比較する項目のみ値が入っているEntityクラス\n");
		sb.append("\t * @return 更新件数\n");
		sb.append("\t */\n");
		sb.append("\t int updateSelective(@Param(\"record\") " + convertCamelCase(tblNm, true)
				+ "Entity record, @Param(\"where\") " + convertCamelCase(tblNm, true) + "Entity where);\n");
		sb.append("}\n");
		return sb.toString();
	}

	/**
	 * 内容作成(XML)
	 * 
	 * @param tblNm
	 * @param tblNmJp
	 * @param pkNmSnk
	 * @param readContentsList
	 * @return
	 */
	private static String createContentsXml(String tblNm, String tblNmJp, String pkNmSnk,
			List<ReadContents> readContentsList) {
		StringBuilder sb = new StringBuilder();

		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		// HEADER
		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		sb.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
		sb.append(
				"<!DOCTYPE mapper PUBLIC \"-//mybatis.org//DTD Mapper 3.0//EN\" \"http://mybatis.org/dtd/mybatis-3-mapper.dtd\">\n");
		sb.append("<mapper namespace=\"jp.progmat.ut.common.infra.mapper." + convertCamelCase(tblNm, false) + "."
				+ convertCamelCase(tblNm, true) + "Mapper\">\n");
		sb.append("\n");

		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		// SELECT
		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		sb.append("\t<select id=\"select\" resultType=\"jp.progmat.ut.common.domain.entity."
				+ convertCamelCase(tblNm, false) + "."
				+ convertCamelCase(tblNm, true) + "Entity\">\n");
		sb.append("\t\tSELECT\n");
		sb.append("\t\t<if test=\"distinct\">\n");
		sb.append("\t\tDISTINCT\n");
		sb.append("\t\t</if>\n");
		for (ReadContents data : readContentsList) {
			sb.append("\t\t\t" + data.getColNmSnk() + ",\n");
		}
		sb = new StringBuilder(sb.substring(0, sb.lastIndexOf(",")) + "\n");
		sb.append("\t\tFROM " + tblNm + "\n");
		sb.append("\t\tWHERE\n");
		sb.append("\t\t\t" + pkNmSnk + " IS NOT NULL\n");
		for (ReadContents data : readContentsList) {
			sb.append("\t\t<if test=\"where." + data.getColNmCml() + " != null\">\n");
			if (!data.getColType().equals(ARRAY)) {
				sb.append("\t\t\tAND " + data.getColNmSnk() + " = #{where." + data.getColNmCml() + "}\n");
			} else if (data.getColType().equals(ARRAY)) {
				sb.append("\t\t\tAND #{where." + data.getColNmCml() + "} = ANY(" + data.getColNmSnk() + ")\n");
			}
			sb.append("\t\t</if>\n");
		}
		sb.append("\t\t<if test=\"orderBy != null\">\n");
		sb.append("\t\tORDER BY\n");
		sb.append("\t\t\t ${orderBy}\n");
		sb.append("\t\t</if>\n");
		sb.append("\t</select>\n");
		sb.append("\n");

		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		// SELECT COUNT
		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		sb.append("\t<select id=\"count\" resultType=\"java.lang.Long\">\n");
		sb.append("\t\tSELECT\n");
		sb.append("\t\t<if test=\"distinct\">\n");
		sb.append("\t\tDISTINCT\n");
		sb.append("\t\t</if>\n");
		sb.append("\t\t\tCOUNT(*)\n");
		sb.append("\t\tFROM " + tblNm + "\n");
		sb.append("\t\tWHERE\n");
		sb.append("\t\t\t" + pkNmSnk + " IS NOT NULL\n");
		for (ReadContents data : readContentsList) {
			sb.append("\t\t<if test=\"where." + data.getColNmCml() + " != null\">\n");
			if (!data.getColType().equals(ARRAY)) {
				sb.append("\t\t\tAND " + data.getColNmSnk() + " = #{where." + data.getColNmCml() + "}\n");
			} else if (data.getColType().equals(ARRAY)) {
				sb.append("\t\t\tAND #{where." + data.getColNmCml() + "} = ANY(" + data.getColNmSnk() + ")\n");
			}
			sb.append("\t\t</if>\n");
		}
		sb.append("\t</select>\n");
		sb.append("\n");

		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		// INSERT
		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		sb.append("\t<insert id=\"insert\">\n");
		sb.append("\t\tINSERT INTO " + tblNm + " (\n");
		for (ReadContents data : readContentsList) {
			sb.append("\t\t\t" + data.getColNmSnk() + ",\n");
		}
		sb = new StringBuilder(sb.substring(0, sb.lastIndexOf(",")) + "\n");
		sb.append("\t\t) VALUES (\n");
		for (ReadContents data : readContentsList) {
			if (data.getColType().equals(CHARACTER)) {
				sb.append("\t\t\t#{record." + data.getColNmCml() + ", jdbcType=VARCHAR},\n");
			} else if (data.getColType().equals(INTEGER)) {
				sb.append("\t\t\t#{record." + data.getColNmCml() + ", jdbcType=INTEGER},\n");
			} else if (data.getColType().equals(TIMESTAMP) || data.getColType().equals(DATE)) {
				sb.append("\t\t\t#{record." + data.getColNmCml() + ", jdbcType=TIMESTAMP},\n");
			} else if (data.getColType().equals(BYTEA)) {
				sb.append("\t\t\t#{record." + data.getColNmCml() + ", jdbcType=BINARY},\n");
			} else if (data.getColType().equals(ARRAY)) {
				sb.append("\t\t\t<if test=\"record." + data.getColNmCml() + " == null\">\n");
				sb.append("\t\t\t\tNULL,\n");
				sb.append("\t\t\t</if>\n");
				sb.append("\t\t\t<if test=\"record." + data.getColNmCml() + " != null\">\n");
				sb.append("\t\t\t\t'{\n");
				sb.append("\t\t\t\t<foreach collection=\"record." + data.getColNmCml()
						+ "\" item=\"item\" separator=\",\">\n");
				sb.append("\t\t\t\t\t${item}\n");
				sb.append("\t\t\t\t</foreach>\n");
				sb.append("\t\t\t\t}',\n");
				sb.append("\t\t\t</if>\n");
			}
		}
		sb = new StringBuilder(sb.substring(0, sb.lastIndexOf(",")) + "\n");
		sb.append("\t\t)\n");
		sb.append("\t</insert>\n");
		sb.append("\n");

		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		// INSERT SELECTIVE
		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		sb.append("\t<insert id=\"insertSelective\">\n");
		sb.append("\t\tINSERT INTO " + tblNm + "\n");
		sb.append("\t\t<trim prefix=\"(\" suffix=\") \" suffixOverrides=\",\">\n");
		for (ReadContents data : readContentsList) {
			sb.append("\t\t\t<if test=\"record." + data.getColNmCml() + " != null\">\n");
			sb.append("\t\t\t\t" + data.getColNmSnk() + ",\n");
			sb.append("\t\t\t</if>\n");
		}
		sb.append("\t\t</trim>\n");
		sb.append("\t\t<trim prefix=\"VALUES (\" suffix=\")\" suffixOverrides=\",\">\n");
		for (ReadContents data : readContentsList) {
			sb.append("\t\t\t<if test=\"record." + data.getColNmCml() + " != null\">\n");
			if (data.getColType().equals(CHARACTER)) {
				sb.append("\t\t\t\t#{record." + data.getColNmCml() + ", jdbcType=VARCHAR},\n");
			} else if (data.getColType().equals(INTEGER)) {
				sb.append("\t\t\t\t#{record." + data.getColNmCml() + ", jdbcType=INTEGER},\n");
			} else if (data.getColType().equals(TIMESTAMP) || data.getColType().equals(DATE)) {
				sb.append("\t\t\t\t#{record." + data.getColNmCml() + ", jdbcType=TIMESTAMP},\n");
			} else if (data.getColType().equals(BYTEA)) {
				sb.append("\t\t\t\t#{record." + data.getColNmCml() + ", jdbcType=BINARY},\n");
			} else if (data.getColType().equals(ARRAY)) {
				sb.append("\t\t\t\t'{\n");
				sb.append("\t\t\t\t<foreach collection=\"record." + data.getColNmCml()
						+ "\" item=\"item\" separator=\",\">\n");
				sb.append("\t\t\t\t\t${item}\n");
				sb.append("\t\t\t\t</foreach>\n");
				sb.append("\t\t\t\t}',\n");
			}
			sb.append("\t\t\t</if>\n");
		}
		sb.append("\t\t</trim>\n");
		sb.append("\t</insert>\n");
		sb.append("\n");

		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		// UPDATE
		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		sb.append("\t<update id=\"update\">\n");
		sb.append("\t\tUPDATE " + tblNm + " SET\n");
		for (ReadContents data : readContentsList) {
			if (data.getColType().equals(CHARACTER)) {
				sb.append(
						"\t\t\t" + data.getColNmSnk() + " = #{record." + data.getColNmCml() + ", jdbcType=VARCHAR},\n");
			} else if (data.getColType().equals(INTEGER)) {
				sb.append(
						"\t\t\t" + data.getColNmSnk() + " = #{record." + data.getColNmCml() + ", jdbcType=INTEGER},\n");
			} else if (data.getColType().equals(TIMESTAMP) || data.getColType().equals(DATE)) {
				sb.append(
						"\t\t\t" + data.getColNmSnk() + " = #{record." + data.getColNmCml()
								+ ", jdbcType=TIMESTAMP},\n");
			} else if (data.getColType().equals(BYTEA)) {
				sb.append("\t\t\t" + data.getColNmSnk() + " = #{record." + data.getColNmCml()
						+ ", jdbcType=BINARY},\n");
			} else if (data.getColType().equals(ARRAY)) {
				sb.append("\t\t\t<if test=\"record." + data.getColNmCml() + " == null\">\n");
				sb.append("\t\t\t\t" + data.getColNmSnk() + " IS NULL,\n");
				sb.append("\t\t\t</if>\n");
				sb.append("\t\t\t<if test=\"record." + data.getColNmCml() + " != null\">\n");
				sb.append("\t\t\t\t" + data.getColNmSnk() + " = '{\n");
				sb.append("\t\t\t\t<foreach collection=\"record." + data.getColNmCml()
						+ "\" item=\"item\" separator=\",\">\n");
				sb.append("\t\t\t\t\t${item}\n");
				sb.append("\t\t\t\t</foreach>\n");
				sb.append("\t\t\t\t}',\n");
				sb.append("\t\t\t</if>\n");
			}
		}
		sb = new StringBuilder(sb.substring(0, sb.lastIndexOf(",")) + "\n");
		sb.append("\t\tWHERE\n");
		sb.append("\t\t\t" + pkNmSnk + " IS NOT NULL\n");
		for (ReadContents data : readContentsList) {
			sb.append("\t\t<if test=\"where." + data.getColNmCml() + " != null\">\n");
			if (!data.getColType().equals(ARRAY)) {
				sb.append("\t\t\tAND " + data.getColNmSnk() + " = #{where." + data.getColNmCml() + "}\n");
			} else if (data.getColType().equals(ARRAY)) {
				sb.append("\t\t\tAND #{where." + data.getColNmCml() + "} = ANY(" + data.getColNmSnk() + ")\n");
			}
			sb.append("\t\t</if>\n");
		}
		sb.append("\t</update>\n");
		sb.append("\n");

		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		// UPDATE SELECTIVE
		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		sb.append("\t<update id=\"updateSelective\">\n");
		sb.append("\t\tUPDATE " + tblNm + "\n");
		sb.append("\t\t<set>\n");
		for (ReadContents data : readContentsList) {
			sb.append("\t\t\t<if test=\"record." + data.getColNmCml() + " != null\">\n");
			if (data.getColType().equals(CHARACTER)) {
				sb.append(
						"\t\t\t\t" + data.getColNmSnk() + " = #{record." + data.getColNmCml()
								+ ", jdbcType=VARCHAR},\n");
			} else if (data.getColType().equals(INTEGER)) {
				sb.append(
						"\t\t\t\t" + data.getColNmSnk() + " = #{record." + data.getColNmCml()
								+ ", jdbcType=INTEGER},\n");
			} else if (data.getColType().equals(TIMESTAMP) || data.getColType().equals(DATE)) {
				sb.append("\t\t\t\t" + data.getColNmSnk() + " = #{record." + data.getColNmCml()
						+ ", jdbcType=TIMESTAMP},\n");
			} else if (data.getColType().equals(BYTEA)) {
				sb.append(
						"\t\t\t\t" + data.getColNmSnk() + " = #{record." + data.getColNmCml()
								+ ", jdbcType=BINARY},\n");
			} else if (data.getColType().equals(ARRAY)) {
				sb.append("\t\t\t\t" + data.getColNmSnk() + " = '{\n");
				sb.append("\t\t\t\t<foreach collection=\"record." + data.getColNmCml()
						+ "\" item=\"item\" separator=\",\">\n");
				sb.append("\t\t\t\t\t${item}\n");
				sb.append("\t\t\t\t</foreach>\n");
				sb.append("\t\t\t\t}',\n");
			}
			sb.append("\t\t\t</if>\n");
		}
		sb.append("\t\t</set>\n");
		sb.append("\t\tWHERE\n");
		sb.append("\t\t\t" + pkNmSnk + " IS NOT NULL\n");
		for (

		ReadContents data : readContentsList) {
			sb.append("\t\t<if test=\"where." + data.getColNmCml() + " != null\">\n");
			if (!data.getColType().equals(ARRAY)) {
				sb.append("\t\t\tAND " + data.getColNmSnk() + " = #{where." + data.getColNmCml() + "}\n");
			} else if (data.getColType().equals(ARRAY)) {
				sb.append("\t\t\tAND #{where." + data.getColNmCml() + "} = ANY(" + data.getColNmSnk() + ")\n");
			}
			sb.append("\t\t</if>\n");
		}
		sb.append("\t</update>\n");
		sb.append("\n");

		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		// FOOTER
		// ■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
		sb.append("</mapper>\n");

		return sb.toString();
	}

	private static void lastShow(boolean result, int stopTime) {
		try {
			System.out.println(result ? "★★★ 成功しました ★★★" : "※※※ 失敗しました ※※※");
			Thread.sleep(stopTime);
			return;
		} catch (Exception e) {
			System.out.println("ストップエラー");
			System.out.println(e);
		}
	}
}
