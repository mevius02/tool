package jp.progmat.ut.common.infra.mapper.mstUser;

import java.util.List;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import jp.progmat.ut.common.domain.entity.mstUser.MstUserEntity;

/** ユーザマスタ */
@Mapper
public interface MstUserMapper {

	/**
	 * [検索]ユーザマスタ
	 * @param distinct DISTINCT [TRUE=DISTINCT付与する/FALSE=DISTINCT付与しない]
	 * @param where    WHERE    イコール比較する項目のみ値が入っているEntityクラス
	 * @param orderBy  ORDER BY ORDER BY句の内容
	 * @return 検索結果リスト
	 */
	 List<MstUserEntity> select(@Param("distinct") boolean distinct, @Param("where") MstUserEntity where, @Param("orderBy") String orderBy);

	/**
	 * [検索]ユーザマスタ
	 * @param distinct DISTINCT [TRUE=DISTINCT付与する/FALSE=DISTINCT付与しない]
	 * @param where    WHERE    イコール比較する項目のみ値が入っているEntityクラス
	 * @param orderBy  ORDER BY ORDER BY句の内容
	 * @return 検索結果件数
	 */
	 long count(@Param("distinct") boolean distinct, @Param("where") MstUserEntity where, @Param("orderBy") String orderBy);

	/**
	 * [追加]ユーザマスタ
	 * @param record レコード
	 * @return 追加件数
	 */
	 int insert(@Param("record") MstUserEntity record);

	/**
	 * [追加]ユーザマスタ(選択項目)
	 * <pre>説明：引数[record]内の各変数値がnull以外の項目のみ対象となる</pre>
	 * <pre>つまり、nullで追加、更新ができない</pre>
	 * @param record レコード
	 * @return 追加件数
	 */
	 int insertSelective(@Param("record") MstUserEntity record);

	/**
	 * [更新]ユーザマスタ
	 * @param record レコード
	 * @param where  WHERE イコール比較する項目のみ値が入っているEntityクラス
	 * @return 更新件数
	 */
	 int update(@Param("record") MstUserEntity record, @Param("where") MstUserEntity where);

	/**
	 * [更新]ユーザマスタ(選択項目)
	 * <pre>説明：引数[record]内の各変数値がnull以外の項目のみ対象となる</pre>
	 * <pre>つまり、nullで追加、更新ができない</pre>
	 * @param record レコード
	 * @param where  WHERE イコール比較する項目のみ値が入っているEntityクラス
	 * @return 更新件数
	 */
	 int updateSelective(@Param("record") MstUserEntity record, @Param("where") MstUserEntity where);
}
