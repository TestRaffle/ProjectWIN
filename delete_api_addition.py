# ===== 以下のコードを main.py に追加してください =====
# 追加場所: /api/admin/license/reset-hw の後、/health の前

class DeleteLicenseRequest(BaseModel):
    license_key: str

@app.post("/api/admin/license/delete")
def admin_delete_license(
    req: DeleteLicenseRequest,
    db: Session = Depends(get_db),
    _: None = Depends(_verify_admin)
):
    """管理者: ライセンスを完全に削除"""
    lic = db.query(License).filter(License.license_key == req.license_key).first()
    if not lic:
        raise HTTPException(status_code=404, detail="License not found")
    
    # 関連する支払い情報も削除（外部キー制約がある場合）
    # db.query(Payment).filter(Payment.license_id == lic.id).delete()
    
    db.delete(lic)
    db.commit()
    return {"success": True, "message": f"License {req.license_key} deleted permanently"}


# ===== 一括削除用API（オプション）=====

class BulkDeleteRequest(BaseModel):
    license_keys: list[str]

@app.post("/api/admin/license/delete-bulk")
def admin_delete_licenses_bulk(
    req: BulkDeleteRequest,
    db: Session = Depends(get_db),
    _: None = Depends(_verify_admin)
):
    """管理者: 複数のライセンスを一括削除"""
    deleted = []
    not_found = []
    
    for key in req.license_keys:
        lic = db.query(License).filter(License.license_key == key).first()
        if lic:
            db.delete(lic)
            deleted.append(key)
        else:
            not_found.append(key)
    
    db.commit()
    return {
        "success": True,
        "deleted": deleted,
        "deleted_count": len(deleted),
        "not_found": not_found
    }
