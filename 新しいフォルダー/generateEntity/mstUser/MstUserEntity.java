package jp.progmat.ut.common.domain.entity.mstUser;

import java.util.Date;

import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import lombok.Getter;
import lombok.Setter;

/** ユーザマスタ */
@Entity
@Getter
@Setter
public class MstUserEntity {

	/** ユーザID */
	@Id
	private String userId;
	/** パスワード */
	private String password;
	/** ユーザ名 */
	private String userName;
	/** 追加ユーザID */
	private String insertUserId;
	/** 追加日時 */
	private Date insertDatetime;
	/** 追加ユーザID */
	private String updateUserId;
	/** 追加日時 */
	private Date updateDatetime;
	/** 削除フラグ */
	private String deleteFlg;

	/** コンストラクター */
	public MstUserEntity() {}

	/**
	 * コンストラクター
	 * @param userId         ユーザID
	 * @param password       パスワード
	 * @param userName       ユーザ名
	 * @param insertUserId   追加ユーザID
	 * @param insertDatetime 追加日時
	 * @param updateUserId   追加ユーザID
	 * @param updateDatetime 追加日時
	 * @param deleteFlg      削除フラグ
	 */
	public MstUserEntity(String userId, String password, String userName, String insertUserId, Date insertDatetime, String updateUserId, Date updateDatetime, String deleteFlg) {
		this.userId = userId;
		this.password = password;
		this.userName = userName;
		this.insertUserId = insertUserId;
		this.insertDatetime = insertDatetime;
		this.updateUserId = updateUserId;
		this.updateDatetime = updateDatetime;
		this.deleteFlg = deleteFlg;
	}
}
