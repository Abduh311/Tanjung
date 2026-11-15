# Security Specification — Report Bazar Matahari Citimall Kuala Kapuas

## 1. Data Invariants
1. **Default-Deny Catch-All**: Every path not explicitly matched is denied (`allow read, write: if false;`).
2. **Strict Path ID Validation (`isValidId`)**: Both `{workspaceId}` and `{submissionId}` must match `^[a-zA-Z0-9_\-]+$` and be at most 128 characters.
3. **Workspace Document Invariant (`BazarWorkspace`)**:
   - Only `workspaceId == "main"` or valid IDs matching `data.workspaceId == workspaceId` are allowed.
   - Must strictly contain only `['workspaceId', 'weeklyTablesJson', 'reviewDocsJson', 'submissionLogsJson', 'updatedAtMs', 'updatedByRole']`.
   - Every string field has strict `.size()` bounds (`weeklyTablesJson <= 900000`, `reviewDocsJson <= 500000`, `submissionLogsJson <= 500000`, `updatedByRole in ['admin', 'report_link', 'public_sync']`).
   - On update, `workspaceId` is immutable (`incoming().workspaceId == existing().workspaceId`) and `updatedAtMs >= existing().updatedAtMs`.
4. **Report Submission Invariant (`ReportSubmission`)**:
   - Must strictly contain only `['id', 'timestamp', 'reportType', 'submittedBy', 'weekTitle', 'tanggal', 'shift', 'salesAmount', 'qtyAmount', 'trxAmount', 'summaryNote', 'createdAtMs']`.
   - `data.id == submissionId`, `reportType in ['shift_sales', 'review_monitoring']`, and all numeric and string bounds are enforced.

## 2. The "Dirty Dozen" Payloads
1. **Ghost Field Injection on Workspace**: `{ workspaceId: "main", ..., isAdmin: true }` -> Rejected by `hasOnly`.
2. **Oversized Payload on Workspace**: `weeklyTablesJson` > 900,000 chars -> Rejected by `.size() <= 900000`.
3. **Invalid Role Enum**: `updatedByRole: "hacker"` -> Rejected by enum check.
4. **Workspace ID Mismatch**: Path `/bazarWorkspaces/main` with `data.workspaceId: "other"` -> Rejected by `data.workspaceId == workspaceId`.
5. **Immutable Workspace ID Change on Update**: Changing `workspaceId` during update -> Rejected by `incoming().workspaceId == existing().workspaceId`.
6. **Timestamp Rollback Attack**: Updating `BazarWorkspace` with `updatedAtMs < existing().updatedAtMs` -> Rejected.
7. **ID Poisoning on Submission**: `/reportSubmissions/bad$id!` -> Rejected by `isValidId(submissionId)`.
8. **Ghost Field on Submission**: Adding `extraField: 123` to `/reportSubmissions/log-1` -> Rejected by `hasOnly`.
9. **Invalid Report Type**: `reportType: "unknown"` -> Rejected by `in ['shift_sales', 'review_monitoring']`.
10. **Negative Sales/Qty/Trx**: `salesAmount: -500` -> Rejected by `>= 0`.
11. **Oversized Summary Note**: `summaryNote` > 2000 chars -> Rejected by `.size() <= 2000`.
12. **Unauthorized Collection Delete**: Deleting `/bazarWorkspaces/main` without admin role -> Rejected (`allow delete: if isAdmin()`).
