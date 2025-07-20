# Consolidating Task and Order Options Plan

## Notes

- User is considering whether to store both task and order options in the same Google Sheet, with columns to distinguish between them, to simplify functions and reduce bugs/errors.
- Current system has separate patient-specific and task manager views for tasks and orders, with different functions for each.
- Task and order data are currently stored in separate sheets ("tasks" and "orders" in taskManager, "orderOptions" in DocumentationInterface). No clear unified options sheet.
- User wants to implement structured, hierarchical task options (like orders), which is not yet implemented.
- Orders have some unique features (faxing, attaching visit info, facesheet), but many features (mentions, tracking, updates) are shared between tasks and orders.
- User has minimal existing data, so migration is not a concern; decision is to proceed with consolidation.
- Approval/signoff logic: If a provider enters an order/task, it is automatically approved; if entered by staff, provider approval is required. The schema should reflect this workflow, possibly by renaming or clarifying the `requiresSignoff` field.
- Distinction clarified: The combined options/taxonomy sheet defines categories, templates, and workflow rules for tasks/orders; the actual log/database sheet records each created task/order instance and its status. Both need clear, separate schemas.
- User's current task log schema was reviewed; noted that a "mentioned" field is present, and further fields may be needed for comprehensive tracking and alignment with consolidated approach.
- User is seeking a unified nomenclature (e.g. 'item', 'activity', etc.) to encompass both tasks and orders, since the main difference is provider approval; schema and code should reflect this simplification where possible.
- Decision: Adopt "ActionItem" as the unified term for tasks, orders, and future patient/non-clinical action items. All schemas, code, and documentation should use this nomenclature for clarity and extensibility.
- User is considering if there are any other typical data fields, organizational strategies, or labeling conventions that should be included to help solve for extensibility, clarity, and future use cases in the ActionItem schema.
- Tags: Consider adding a `tags` field to ActionItems for flexible, user-defined categorization (e.g., urgent, billing-related, follow-up).
- Use `defaultUrgency` and `defaultImportance` in the actionItemOptions sheet to drive backend priority calculation, but display only a single `priority` field in the actionItems sheet for assigned users to keep the UI simple and clear.
- Relationship tracking (e.g., parentId, relatedIds) and recurrence fields (e.g., isRecurring, recurrencePattern) should be implemented in the actionItems sheet, not the options sheet. The options sheet may include default recurrence settings (e.g., defaultRecurrence, allowsRecurrence) for template purposes.
- Options table structure: Consider the trade-offs between one row per specific option (cleaner for unique selections, e.g., left/right) versus storing multiple selectable options as an array/list in a single column (better for multi-select categories like patient education topics). UI/UX and processing code should inform the final structure.
- Template support: Add `isTemplate` (boolean) and `templateId` (reference) columns to both actionItemOptions and actionItems sheets as needed. Templates can pre-populate fields for common actions and support efficient creation of recurring or standardized ActionItems.
- ActionItem, ActionItemOptions, and ActionItemsAudit schemas finalized; ready for implementation in Google Apps Script as per new unified plan.
- Google Apps Script project for ActionItems created and ready for code implementation.
- Initial ActionItems CRUD code and file structure created in Apps Script project; implementation of hierarchical options and integration features is next.

## Task List

- [x] Review current implementation of task and order options storage.

- [x] Evaluate pros and cons of consolidating into a single Google Sheet.
- [x] Propose a schema for combined task/order options sheet (with distinguishing columns and support for hierarchy in both, and clarify approval/signoff logic).
  - [x] Specify column headers for options/taxonomy sheet.
  - [x] Specify column headers for task/order log/database sheet (review and improve based on current schema, ensure fields like "mentioned" and others are included).
  - [x] Consider and implement a tagging system for ActionItems if desired.
  - [x] Consider and implement relationship tracking (parentId, relatedIds) and recurrence fields (isRecurring, recurrencePattern, etc.) in the actionItems schema.
  - [x] Consider and implement template support (`isTemplate`, `templateId`) in both options and actionItems schemas.
- [x] Create initial CRUD code for ActionItems in Google Apps Script.
- [ ] Design and implement hierarchical task options similar to order options.
- [ ] Implement new Google Apps Script for ActionItems notifications, comments, mentions, and integration with DocumentationInterface.
  - [x] Set up new Google Apps Script project (create script, clasp, link to GitHub)
  - [x] Create initial Code.js and ActionItemCRUD.js with core CRUD logic
- [ ] Assess impact on existing code (taskManager, DocumentationInterface, etc.).
- [ ] Recommend next steps based on findings.

## Current Goal

Continue ActionItems Apps Script implementation: options, notifications, integration.
