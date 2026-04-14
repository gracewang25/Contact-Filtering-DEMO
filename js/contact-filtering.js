// contact-filtering.js
// Upload contacts, detect likely duplicates, edit entries inline and manage linked interactions in QuickBase.
// "use strict";


(function () {
	// Configuration
	const QB_CONFIG = {
		realm: "demo.quickbase.com",
		appId: "APP_DEMO_001",
		contactTableId: "TABLE_CONTACTS_DEMO",
		affiliationTableId: "TABLE_AFFILIATIONS_DEMO",
		interactionTableId: "TABLE_INTERACTIONS_DEMO",
		contactFields: {
			recordId: 101,
			firstName: 102,
			lastName: 103,
			email: 104,
			relatedAffiliation: 105
		},
		affiliationFields: {
			recordId: 201,
			affiliationName: 202
		},
		interactionFields: {
			recordId: 301,
			showName: 302,
			relatedContact: 303,
			type: 304,
			source: 305,
			sourceDate: 306,
		}
	};


	// Demo mode — data and config live in contact-filter.demo.js
	const Demo = window.ContactFilteringDemo;
	const USE_DEMO_CONTACTS = !!(Demo && Demo.useDemo);

	const DEFAULT_INTERACTION_TYPE = "In-Person Meeting";
	const DUPLICATE_SCORE_THRESHOLD = 74;
	
	// UI update delay
	const SUCCESS_RESET_MS = 2600;

	const appToken = "DEMO_APP_TOKEN";

	// Runtime state
	const authStateByTable = {};
	let storedContacts = [];
	let allSheetObjs = [];
	let currentMatches = [];
	let nonMatches = [];
	let sheetObjByRow = new Map();
	let matchByRow = new Map();
	let interactionsByContactId = new Map();
	let uploadScopedSourceDate = "";
	let contactErrorMsg = "";
	let interactionErrorMsg = "";

	// DOM elements
	const fileInput = document.getElementById("fileInput");
	const processBtn = document.getElementById("processBtn");
	const resultsContainer = document.getElementById("results");
	const duplicatesList = document.getElementById("duplicatesList");
	const newContactsList = document.getElementById("newContactsList");
	const contactStatus = document.getElementById("contactStatus");
	const duplicateCount = document.getElementById("duplicateCount");
	const newContactCount = document.getElementById("newContactCount");
	const totalCount = document.getElementById("totalCount");
	const exportResultsBtn = document.getElementById("exportResultsBtn");
	const updateAllBtn = document.getElementById("updateAllBtn");
	const addAllBtn = document.getElementById("addAllBtn");
	const interactionShowName = document.getElementById("interactionShowName");
	const uploadSourceDate = document.getElementById("uploadSourceDate");
	const bulkInteractionType = document.getElementById("bulkInteractionType");
	const interactionSource = document.getElementById("interactionSource");
	const addInteractionAllBtn = document.getElementById("addInteractionAllBtn");
	const removeInteractionAllBtn = document.getElementById("removeInteractionAllBtn");
	const viewInteractionAllBtn = document.getElementById("viewInteractionAllBtn");
	const interactionStatus = document.getElementById("interactionStatus");
	const interactionOverview = document.getElementById("interactionOverview");
	const interactionModal = document.getElementById("interactionModal");
	const interactionModalOverlay = document.getElementById("interactionModalOverlay");
	const interactionModalCloseBtn = document.getElementById("interactionModalCloseBtn");
	const interactionModalCancelBtn = document.getElementById("interactionModalCancelBtn");
	const interactionModalTitle = document.getElementById("interactionModalTitle");
	const interactionModalContext = document.getElementById("interactionModalContext");
	const interactionForm = document.getElementById("interactionForm");
	const interactionFieldShowName = document.getElementById("interactionFieldShowName");
	const interactionFieldContactId = document.getElementById("interactionFieldContactId");
	const interactionFieldType = document.getElementById("interactionFieldType");
	const interactionFieldSource = document.getElementById("interactionFieldSource");
	const interactionModalSubmitBtn = document.getElementById("interactionModalSubmitBtn");
	const errorBanner = document.getElementById("errorBanner");
	const loadDemoBtn = document.getElementById("loadDemoBtn");

	let interactionModalState = null;
	const CFUtils = window.ContactFilteringUtils;
	const CFData = window.ContactFilteringData;

	if (!CFUtils || !CFData) {
		throw new Error("Missing required script dependencies: contact-filtering.utils.js and contact-filtering.data.js must load before contact-filtering.js.");
	}

	// Utility helpers
	function normalizeForMatch(v) {
		return CFUtils.normalizeForMatch(v);
	}

	function normalizeForDisplay(v) {
		return CFUtils.normalizeForDisplay(v);
	}

	function escapeHtml(s) {
		return CFUtils.escapeHtml(s);
	}
	
	function chunkArray(items, chunkSize) {
		return CFUtils.chunkArray(items, chunkSize);
	}

	// Status and UI helpers
	function syncErrorBanner() {
		const messages = [contactErrorMsg, interactionErrorMsg].filter(Boolean);
		if (messages.length) {
			errorBanner.innerHTML = "";
			messages.forEach(function (m) {
				const div = document.createElement("div");
				div.textContent = m;
				errorBanner.appendChild(div);
			});
			errorBanner.hidden = false;
			document.body.style.paddingTop = errorBanner.offsetHeight + "px";
		} else {
			errorBanner.hidden = true;
			document.body.style.paddingTop = "";
		}
	}

	function showContactStatus(msg, isError) {
		contactStatus.innerHTML = isError ? "" : `<div class="success">${msg}</div>`;
		contactErrorMsg = isError ? msg : "";
		syncErrorBanner();
	}

	function showInteractionStatus(msg, isError) {
		interactionStatus.innerHTML = `<div class="${isError ? "error" : "success"}">${msg}</div>`;
		interactionErrorMsg = isError ? msg : "";
		syncErrorBanner();
	}

	function qbContactRecordUrl(recordId) {
		if (!recordId) return "";
		return `https://${QB_CONFIG.realm}/db/${QB_CONFIG.contactTableId}?a=dr&rid=${Number(recordId)}`;
	}

	function getHeadersForLog(headers) {
		return CFUtils.getHeadersForLog(headers);
	}

	function getAjaxErrorMessage(err) {
		return CFUtils.getAjaxErrorMessage(err);
	}

	function assertNoLineErrors(response, actionLabel) {
		CFUtils.assertNoLineErrors(response, actionLabel);
	}

	function resolveTokenExpiryMs(data) {
		return CFUtils.resolveTokenExpiryMs(data);
	}

	// QB Auth
	function getTempAuth(realm, dbid, appTokenValue) {
		return new Promise(function (resolve, reject) {
			const headers = {
				"QB-Realm-Hostname": realm,
				userAgent: "QB APIGateway"
			};

			if (appTokenValue) {
				headers["QB-App-Token"] = appTokenValue;
			}

			$.ajax({
				url: `https://api.quickbase.com/v1/auth/temporary/${dbid}`,
				method: "GET",
				headers: headers,
				xhrFields: { withCredentials: true },
				success: function (data) {
					console.info("[QB] Temporary auth token acquired", {
						tableId: dbid,
						requestHeaders: getHeadersForLog(headers),
						expiresAt: data && (data.expiresAt || data.expiration || data.temporaryAuthorizationExpiration || null)
					});
					resolve({
						headers: {
							"QB-Realm-Hostname": realm,
							userAgent: "QB APIGateway",
							Authorization: `QB-TEMP-TOKEN ${data.temporaryAuthorization}`
						},
						expiresAtMs: resolveTokenExpiryMs(data)
					});
				},
				error: function (xhr, _status, error) {
					console.error("[QB] Temporary auth failed", {
						tableId: dbid,
						status: xhr && xhr.status,
						error: getAjaxErrorMessage(xhr || error)
					});
					reject(new Error(`Auth failed: ${error}`));
				}
			});
		});
	}

	async function getAuthHeadersForTable(tableId, forceRefresh) {
		const key = String(tableId || QB_CONFIG.contactTableId);
		const now = Date.now();
		const existing = authStateByTable[key];

		if (!forceRefresh && existing && existing.headers && existing.expiresAtMs && now < existing.expiresAtMs) {
			return Promise.resolve(existing.headers);
		}

		const authResult = await getTempAuth(QB_CONFIG.realm, key, appToken);
		authStateByTable[key] = {
			headers: authResult.headers,
			expiresAtMs: authResult.expiresAtMs
		};
		return authResult.headers;
	}

	// QB API Calls

	async function qbRequest(url, method, payload, tableId) {
		if (USE_DEMO_CONTACTS) {
			return Demo.qbRequest(url, method, payload, QB_CONFIG);
		}

		const targetTableId = tableId || QB_CONFIG.contactTableId;

		// Centralized API calls with token refresh handling and retry
		function runWithHeaders(headers) {
			console.info("[QB] Request", {
				url,
				method,
				tableId: targetTableId,
				headers: getHeadersForLog(headers),
				payload
			});
			return $.ajax({
				url: url,
				method: method,
				headers: headers,
				dataType: "json",
				contentType: "application/json; charset=utf-8",
				data: payload ? JSON.stringify(payload) : undefined
			});
		}

		try {
			const headers = await getAuthHeadersForTable(targetTableId, false);
			const response = await runWithHeaders(headers);
			console.info("[QB] Response success", { url, method, tableId: targetTableId, response });
			return response;
		} catch (xhr) {
			if (xhr && xhr.status === 401) {
				console.warn("[QB] 401 received. Refreshing token and retrying.", { url, tableId: targetTableId });
				try {
					const refreshedHeaders = await getAuthHeadersForTable(targetTableId, true);
					const response = await runWithHeaders(refreshedHeaders);
					console.info("[QB] Retry success", { url, method, tableId: targetTableId, response });
					return response;
				} catch (retryErr) {
					console.error("[QB] Retry failed", {
						url,
						method,
						tableId: targetTableId,
						status: retryErr && retryErr.status,
						error: getAjaxErrorMessage(retryErr)
					});
					throw retryErr;
				}
			}

			console.error("[QB] Request failed", {
				url,
				method,
				tableId: targetTableId,
				status: xhr && xhr.status,
				error: getAjaxErrorMessage(xhr)
			});
			throw xhr;
		}
	}

	// Contact loading and enrichment
	async function loadAffiliationNameMap(affiliationRecordIds) {
		if (!affiliationRecordIds.length) {
			return Promise.resolve(new Map());
		}

		async function queryAffiliationRowsByIds(ids) {
			const where = ids
				.map((id) => `{${QB_CONFIG.affiliationFields.recordId}.EX.${id}}`)
				.join("OR");

			const response = await qbRequest("https://api.quickbase.com/v1/records/query", "POST", {
				from: QB_CONFIG.affiliationTableId,
				select: [QB_CONFIG.affiliationFields.recordId, QB_CONFIG.affiliationFields.affiliationName],
				where
			}, QB_CONFIG.affiliationTableId);
			return response && response.data ? response.data : [];
		}

		async function fetchIdsWithSplit(ids) {
			if (!ids.length) return Promise.resolve([]);

			try {
				return await queryAffiliationRowsByIds(ids);
			} catch (err) {
				if (ids.length === 1) {
					console.warn("[QB] Skipping invalid/unqueryable affiliation id", {
						affiliationId: ids[0],
						error: getAjaxErrorMessage(err)
					});
					return [];
				}

				const mid = Math.ceil(ids.length / 2);
				const left = ids.slice(0, mid);
				const right = ids.slice(mid);
				const parts = await Promise.all([fetchIdsWithSplit(left), fetchIdsWithSplit(right)]);
				return parts[0].concat(parts[1]);
			}
		}

		const idChunks = chunkArray(affiliationRecordIds, 80);
		const requests = idChunks.map((ids) => fetchIdsWithSplit(ids));

		const results = await Promise.all(requests);
		const affiliationNameByRecordId = new Map();
		results.flat().forEach((row) => {
			const id = Number(row[QB_CONFIG.affiliationFields.recordId]?.value || 0);
			const name = normalizeForDisplay(row[QB_CONFIG.affiliationFields.affiliationName]?.value || "");
			if (id) affiliationNameByRecordId.set(id, name);
		});
		return affiliationNameByRecordId;
	}

	async function loadContactsFromQB() {
		if (USE_DEMO_CONTACTS) {
			storedContacts = Demo.existingContacts.map((contact) => {
				const firstName = normalizeForDisplay(contact.firstName || "");
				const lastName = normalizeForDisplay(contact.lastName || "");
				const name = [firstName, lastName].filter(Boolean).join(" ");
				const email = normalizeForDisplay(contact.email || "");
				const affiliation = normalizeForDisplay(contact.affiliation || "");
				return {
					recordId: Number(contact.recordId || 0),
					firstName,
					lastName,
					name,
					email,
					affiliation,
					nameNorm: normalizeForMatch(name),
					emailNorm: normalizeForMatch(email),
					affiliationNorm: normalizeForMatch(affiliation),
					combinedNorm: normalizeForMatch(`${name} ${email} ${affiliation}`)
				};
			}).filter((contact) => contact.recordId && (contact.name || contact.email || contact.affiliation));

			showContactStatus(`Ready. Loaded ${storedContacts.length} demo contacts for matching.`, false);
			processBtn.disabled = false;
			return;
		}

		showContactStatus("Loading contacts from QuickBase...", false);

		const response = await qbRequest("https://api.quickbase.com/v1/records/query", "POST", {
			from: QB_CONFIG.contactTableId,
			select: [
				QB_CONFIG.contactFields.recordId,
				QB_CONFIG.contactFields.firstName,
				QB_CONFIG.contactFields.lastName,
				QB_CONFIG.contactFields.email,
				QB_CONFIG.contactFields.relatedAffiliation
			]
		}, QB_CONFIG.contactTableId);

		// Existing contacts from QuickBase
		const qbRows = response && response.data ? response.data : [];
		const affiliationRecordIds = Array.from(new Set(
			qbRows
				.map((row) => Number(row[QB_CONFIG.contactFields.relatedAffiliation]?.value || 0))
				.filter(Boolean)
		));

		let affiliationNameByRecordId = new Map();
		try {
			affiliationNameByRecordId = await loadAffiliationNameMap(affiliationRecordIds);
		} catch (err) {
			console.warn("[QB] Affiliation lookup failed; continuing without affiliation enrichment", {
				error: getAjaxErrorMessage(err)
			});
		}

		storedContacts = qbRows
			.map((row) => {
				const firstName = normalizeForDisplay(row[QB_CONFIG.contactFields.firstName]?.value || "");
				const lastName = normalizeForDisplay(row[QB_CONFIG.contactFields.lastName]?.value || "");
				const name = [firstName, lastName].filter(Boolean).join(" ");
				const email = normalizeForDisplay(row[QB_CONFIG.contactFields.email]?.value || "");
				const recordId = Number(row[QB_CONFIG.contactFields.recordId]?.value || 0);
				const relatedAffiliationId = Number(row[QB_CONFIG.contactFields.relatedAffiliation]?.value || 0);
				const affiliation = normalizeForDisplay(affiliationNameByRecordId.get(relatedAffiliationId) || "");
				return {
					recordId,
					firstName,
					lastName,
					name,
					email,
					affiliation,
					nameNorm: normalizeForMatch(name),
					emailNorm: normalizeForMatch(email),
					affiliationNorm: normalizeForMatch(affiliation),
					combinedNorm: normalizeForMatch(`${name} ${email} ${affiliation}`)
				};
			})
			.filter((contact) => contact.recordId && (contact.name || contact.email || contact.affiliation));

		showContactStatus(`Ready. Loaded ${storedContacts.length} contacts from QuickBase.`, false);
		processBtn.disabled = false;
	}

	// Data helpers

	// Injects built-in sample contacts as if a file were uploaded — used in demo mode
	function loadDemoUpload() {
		allSheetObjs = Demo.uploadContacts.map(function (c, i) {
			const firstName = normalizeForDisplay(c.firstName || "");
			const lastName = normalizeForDisplay(c.lastName || "");
			const name = [firstName, lastName].filter(Boolean).join(" ");
			const email = normalizeForDisplay(c.email || "");
			const affiliation = normalizeForDisplay(c.affiliation || "");
			return {
				rowIndex: i + 2,
				firstName, lastName, name, email, affiliation,
				nameNorm: normalizeForMatch(name),
				emailNorm: normalizeForMatch(email),
				affiliationNorm: normalizeForMatch(affiliation),
				combinedNorm: normalizeForMatch(`${name} ${email} ${affiliation}`)
			};
		});

		rebuildSheetObjectLookup();
		uploadScopedSourceDate = "";
		currentMatches = searchMatches(allSheetObjs);
		interactionsByContactId = new Map();
		interactionOverview.style.display = "none";
		interactionOverview.innerHTML = "";
		renderResults();
		showContactStatus(`Processed ${allSheetObjs.length} demo contacts from sample upload.`, false);
		showInteractionStatus("Interaction tools are ready.", false);

		const label = document.querySelector(".file-label");
		if (label) {
			label.innerHTML = `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
				<path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/>
			</svg> demo-upload.xlsx`;
		}
	}

	function parseWorkbook(file) {
		return CFData.parseWorkbook(file);
	}

	function sheetRowsToObjects(rows) {
		return CFData.sheetRowsToObjects(rows);
	}

	function searchMatches(sheetObjs) {
		const result = CFData.searchMatches(sheetObjs, storedContacts, DUPLICATE_SCORE_THRESHOLD);
		nonMatches = result.nonMatches;
		matchByRow = result.matchByRow;
		return result.matches;
	}

	// Row and match lookup helpers

	function rebuildSheetObjectLookup() {
		sheetObjByRow = new Map(allSheetObjs.map((obj) => [Number(obj.rowIndex), obj]));
	}

	function rebuildMatchLookup() {
		matchByRow = new Map(currentMatches.map((match) => [Number(match.item.rowIndex), match]));
	}

	function getObjectByRow(row) {
		return sheetObjByRow.get(Number(row)) || null;
	}

	function getMatchByRow(row) {
		return matchByRow.get(Number(row)) || null;
	}

	// Interaction form value helpers

	function getInteractionTypeValue() {
		const selected = normalizeForDisplay(bulkInteractionType ? bulkInteractionType.value : "");
		return selected || DEFAULT_INTERACTION_TYPE;
	}

	function getShowNameValue() {
		return normalizeForDisplay(interactionShowName ? interactionShowName.value : "");
	}

	function syncShowNameSourceOption(selectEl, showName) {
		if (!selectEl) return;
		const firstOption = selectEl.options && selectEl.options[0];
		if (!firstOption) return;

		const previousDynamicValue = firstOption.dataset.dynamicValue || "";
		const currentValue = normalizeForDisplay(selectEl.value);

		firstOption.dataset.dynamicValue = showName;
		firstOption.value = showName;
		firstOption.textContent = showName || "Show Name (set above)";

		if (!currentValue || currentValue === previousDynamicValue) {
			selectEl.value = showName;
		}
	}

	function syncSourceControlsFromShowName() {
		const showName = getShowNameValue();
		syncShowNameSourceOption(interactionSource, showName);
		syncShowNameSourceOption(interactionFieldSource, showName);

		if (interactionFieldShowName) {
			interactionFieldShowName.value = showName;
		}
	}

	function getInteractionSourceValue(showName, explicitSource) {
		const sourceFromInput = normalizeForDisplay(explicitSource);
		return sourceFromInput || showName;
	}

	// Date helpers

	function getTodayIsoDate() {
		return CFUtils.getTodayIsoDate();
	}

	function normalizeDateForQuickBase(rawDate) {
		return CFUtils.normalizeDateForQuickBase(rawDate);
	}

	function getUploadScopedSourceDateValue(sourceDateInput) {
		if (!uploadScopedSourceDate) {
			uploadScopedSourceDate = normalizeDateForQuickBase(sourceDateInput);
		}
		return uploadScopedSourceDate;
	}

	// Interaction modal management

	function closeInteractionModal() {
		if (!interactionModal) return;
		interactionModal.hidden = true;
		interactionModalState = null;
		if (interactionForm) {
			interactionForm.reset();
		}
		if (interactionFieldContactId) {
			interactionFieldContactId.dataset.contactId = "";
		}
		if (interactionFieldSource) {
			syncShowNameSourceOption(interactionFieldSource, getShowNameValue());
		}
		if (interactionFieldShowName && interactionShowName) {
			interactionFieldShowName.value = normalizeForDisplay(interactionShowName.value);
		}
		if (interactionFieldType) {
			interactionFieldType.value = getInteractionTypeValue();
		}
	}

	function openInteractionModalForRow(row) {
		const item = getObjectByRow(row);
		if (!item) return;

		const match = getMatchByRow(row);
		const linkedRecordId = Number(item.linkedRecordId || 0);
		const contactDisplayName = normalizeForDisplay(item.name || (match && match.by && match.by.name) || "");
		interactionModalState = {
			mode: "single",
			row: Number(row)
		};

		interactionModalTitle.textContent = "Add Interaction";
		interactionModalContext.textContent = `Row ${item.rowIndex}: ${item.name || "Unnamed Contact"}`;
		interactionFieldShowName.value = normalizeForDisplay(interactionShowName ? interactionShowName.value : "");
		interactionFieldContactId.value = contactDisplayName || "Unnamed Contact";
		interactionFieldContactId.dataset.contactId = linkedRecordId ? String(linkedRecordId) : "";
		interactionFieldType.value = getInteractionTypeValue();
		syncShowNameSourceOption(interactionFieldSource, getShowNameValue());
		interactionModal.hidden = false;
	}

	// Contact linking

	async function ensureLinkedContactRecordForRow(row) {
		const item = getObjectByRow(row);
		if (!item) {
			throw new Error("Contact row not found.");
		}

		if (item.linkedRecordId) {
			return Promise.resolve(Number(item.linkedRecordId));
		}

		const match = getMatchByRow(row);
		if (match && match.by && match.by.recordId) {
			item.linkedRecordId = Number(match.by.recordId);
			return Promise.resolve(Number(item.linkedRecordId));
		}

		const ids = await createContactsForItems([item]);
		const createdId = Number(item.linkedRecordId || ids[0] || 0);
		if (!createdId) {
			throw new Error("Unable to create/link contact before adding interaction.");
		}
		item.linkedRecordId = createdId;
		return createdId;
	}

	// Button state helpers

	function setButtonLoading(btn, loadingText) {
		btn.disabled = true;
		btn.dataset.originalText = btn.dataset.originalText || btn.textContent;
		btn.textContent = loadingText;
	}

	function resetButton(btn) {
		btn.disabled = false;
		if (btn.dataset.originalText) {
			btn.textContent = btn.dataset.originalText;
		}
	}

	function markButtonSuccess(btn, text) {
		btn.disabled = true;
		btn.textContent = text;
		btn.classList.add("btn-success");
		setTimeout(() => {
			btn.classList.remove("btn-success");
			resetButton(btn);
		}, SUCCESS_RESET_MS);
	}

	// Contact payload
	function contactToRecordPayload(contact) {
		return {
			[QB_CONFIG.contactFields.firstName]: { value: normalizeForDisplay(contact.firstName) },
			[QB_CONFIG.contactFields.lastName]: { value: normalizeForDisplay(contact.lastName) },
			[QB_CONFIG.contactFields.email]: { value: normalizeForDisplay(contact.email) }
		};
	}

	function getEditableContactFromCard(card, fallbackContact) {
		if (!card || !fallbackContact) return fallbackContact;

		const firstNameInput = card.querySelector(".edit-first-name");
		const lastNameInput = card.querySelector(".edit-last-name");
		const emailInput = card.querySelector(".edit-email");

		if (!firstNameInput || !lastNameInput || !emailInput) {
			return fallbackContact;
		}

		const firstName = normalizeForDisplay(firstNameInput.value);
		const lastName = normalizeForDisplay(lastNameInput.value);
		const email = normalizeForDisplay(emailInput.value);
		const name = [firstName, lastName].filter(Boolean).join(" ");

		return {
			rowIndex: fallbackContact.rowIndex,
			firstName,
			lastName,
			email,
			affiliation: fallbackContact.affiliation,
			name,
			nameNorm: normalizeForMatch(name),
			emailNorm: normalizeForMatch(email),
			affiliationNorm: normalizeForMatch(fallbackContact.affiliation),
			combinedNorm: normalizeForMatch(`${name} ${email} ${fallbackContact.affiliation || ""}`),
			linkedRecordId: fallbackContact.linkedRecordId,
			raw: fallbackContact.raw
		};
	}

	// Contact write operations

	async function updateContact(row, recordId, btn, card) {
		const contact = getObjectByRow(row);
		if (!contact || !recordId) return Promise.resolve(null);
		const editedContact = getEditableContactFromCard(card, contact);

		// Keep in-memory row synced so later actions use what user actually typed.
		Object.assign(contact, editedContact);

		setButtonLoading(btn, "Updating...");

		try {
			const response = await qbRequest("https://api.quickbase.com/v1/records", "POST", {
				to: QB_CONFIG.contactTableId,
				data: [Object.assign({ [QB_CONFIG.contactFields.recordId]: { value: Number(recordId) } }, contactToRecordPayload(editedContact))]
			}, QB_CONFIG.contactTableId);

			assertNoLineErrors(response, "Update contact");
			const metadata = response && response.metadata ? response.metadata : {};
			const updatedRecordIds = metadata.updatedRecordIds || [];
			const unchangedRecordIds = metadata.unchangedRecordIds || [];
			const numericRecordId = Number(recordId);

			if (updatedRecordIds.includes(numericRecordId)) {
				markButtonSuccess(btn, "Updated");
				showContactStatus(`Updated contact for row ${contact.rowIndex}.`, false);
			} else if (unchangedRecordIds.includes(numericRecordId)) {
				resetButton(btn);
				showContactStatus(`No field changes detected for row ${contact.rowIndex}; QuickBase left the record unchanged.`, false);
			} else {
				resetButton(btn);
				showContactStatus(`Update request completed for row ${contact.rowIndex}; verify field permissions/mapping if values did not change.`, false);
			}

			renderResults();
		} catch (error) {
			console.error("[Contact Update] Failed", {
				row,
				recordId,
				payload: contactToRecordPayload(editedContact),
				error: getAjaxErrorMessage(error)
			});
			resetButton(btn);
			btn.textContent = "Retry Update";
			showContactStatus(`Update failed: ${getAjaxErrorMessage(error)}`, true);
		}
	}

	async function updateAllContacts() {
		if (currentMatches.length === 0) return;
		setButtonLoading(updateAllBtn, `Updating ${currentMatches.length} contacts...`);

		const data = currentMatches.map((m) => {
			const payload = contactToRecordPayload(m.item);
			payload[QB_CONFIG.contactFields.recordId] = { value: Number(m.by.recordId) };
			return payload;
		});

		try {
			const response = await qbRequest("https://api.quickbase.com/v1/records", "POST", {
				to: QB_CONFIG.contactTableId,
				data
			}, QB_CONFIG.contactTableId);

			assertNoLineErrors(response, "Bulk update contacts");
			const metadata = response && response.metadata ? response.metadata : {};
			const updatedCount = (metadata.updatedRecordIds || []).length;
			const unchangedCount = (metadata.unchangedRecordIds || []).length;

			if (updatedCount > 0) {
				markButtonSuccess(updateAllBtn, `Updated ${updatedCount}`);
				document.querySelectorAll(".btn-update").forEach((button) => markButtonSuccess(button, "Updated"));
			}

			if (updatedCount === 0 && unchangedCount > 0) {
				resetButton(updateAllBtn);
				showContactStatus(`No bulk updates were applied. ${unchangedCount} records were unchanged.`, false);
			} else {
				showContactStatus(`Bulk update complete. Updated: ${updatedCount}, unchanged: ${unchangedCount}.`, false);
			}
		} catch (error) {
			console.error("[Bulk Contact Update] Failed", {
				count: currentMatches.length,
				error: getAjaxErrorMessage(error)
			});
			resetButton(updateAllBtn);
			updateAllBtn.textContent = "Retry Update All";
			showContactStatus(`Bulk update failed: ${getAjaxErrorMessage(error)}`, true);
		}
	}

	// Contact creation

	async function createContactsForItems(items) {
		if (!items.length) return Promise.resolve([]);

		const payload = {
			to: QB_CONFIG.contactTableId,
			data: items.map((item) => contactToRecordPayload(item))
		};

		const response = await qbRequest("https://api.quickbase.com/v1/records", "POST", payload, QB_CONFIG.contactTableId);
		assertNoLineErrors(response, "Create contacts");
		const createdIds = (response && response.metadata && response.metadata.createdRecordIds) || [];
		if (!createdIds.length && items.length) {
			throw new Error("No contact records were created.");
		}

		items.forEach((item, index) => {
			const rid = Number(createdIds[index] || 0);
			if (rid) {
				item.linkedRecordId = rid;
			}
		});
		return createdIds.map((id) => Number(id));
	}

	async function createContact(row, btn) {
		const contact = getObjectByRow(row);
		if (!contact) return Promise.resolve(null);

		setButtonLoading(btn, "Creating...");

		try {
			const ids = await createContactsForItems([contact]);
			if (ids.length) {
				markButtonSuccess(btn, "Created");
				showContactStatus(`Created contact for row ${contact.rowIndex}.`, false);
				renderResults();
			}
		} catch (error) {
			resetButton(btn);
			btn.textContent = "Retry Create";
			showContactStatus(`Create failed: ${error.message || error}`, true);
		}
	}

	async function createAllContacts() {
		if (nonMatches.length === 0) return;

		setButtonLoading(addAllBtn, `Adding ${nonMatches.length} contacts...`);

		try {
			const createdIds = await createContactsForItems(nonMatches);
			markButtonSuccess(addAllBtn, `Added ${createdIds.length}`);
			document.querySelectorAll(".btn-create").forEach((btn) => markButtonSuccess(btn, "Created"));
			showContactStatus(`Added ${createdIds.length} new contacts.`, false);
			renderResults();
		} catch (error) {
			resetButton(addAllBtn);
			addAllBtn.textContent = "Retry Add All";
			showContactStatus(`Bulk create failed: ${error.message || error}`, true);
		}
	}

	// Contact ID resolution

	async function ensureContactIdsForAllUploaded(createMissing) {
		allSheetObjs.forEach((item) => {
			if (!item.linkedRecordId) {
				const match = getMatchByRow(item.rowIndex);
				if (match) {
					item.linkedRecordId = Number(match.by.recordId);
				}
			}
		});

		const missing = allSheetObjs.filter((item) => !item.linkedRecordId);
		if (!missing.length || !createMissing) {
			return Promise.resolve(allSheetObjs.map((item) => Number(item.linkedRecordId)).filter(Boolean));
		}

		await createContactsForItems(missing);
		return allSheetObjs.map((item) => Number(item.linkedRecordId)).filter(Boolean);
	}

	// Interaction QB operations

	async function createInteractionsForContactIds(contactIds, interactionType, showNameValue, sourceValue, sourceDateValue) {
		if (!contactIds.length) return Promise.resolve([]);
		const showName = normalizeForDisplay(showNameValue);
		const source = getInteractionSourceValue(showName, sourceValue);
		const sourceDate = normalizeForDisplay(sourceDateValue);

		const payload = {
			to: QB_CONFIG.interactionTableId,
			data: contactIds.map((contactId) => ({
				[QB_CONFIG.interactionFields.showName]: { value: showName },
				[QB_CONFIG.interactionFields.relatedContact]: { value: Number(contactId) },
				[QB_CONFIG.interactionFields.type]: { value: interactionType },
				[QB_CONFIG.interactionFields.source]: { value: source },
				[QB_CONFIG.interactionFields.sourceDate]: { value: sourceDate }
			}))
		};

		const response = await qbRequest("https://api.quickbase.com/v1/records", "POST", payload, QB_CONFIG.interactionTableId);
		assertNoLineErrors(response, "Create interactions");
		const createdIds = (response && response.metadata && response.metadata.createdRecordIds) || [];
		return createdIds.map((id) => Number(id));
	}

	function buildWhereForContactIds(contactIds) {
		const uniqueIds = Array.from(new Set(contactIds.map((id) => Number(id)).filter(Boolean)));
		if (!uniqueIds.length) return "";
		return uniqueIds.map((id) => `{${QB_CONFIG.interactionFields.relatedContact}.EX.${id}}`).join("OR");
	}

	async function queryExistingInteractions(contactIds) {
		const where = buildWhereForContactIds(contactIds);
		if (!where) return Promise.resolve([]);

		const response = await qbRequest("https://api.quickbase.com/v1/records/query", "POST", {
			from: QB_CONFIG.interactionTableId,
			select: [
				QB_CONFIG.interactionFields.recordId,
				QB_CONFIG.interactionFields.relatedContact,
				QB_CONFIG.interactionFields.type,
				QB_CONFIG.interactionFields.source
			],
			where
		}, QB_CONFIG.interactionTableId);
		return response && response.data ? response.data : [];
	}

	// Interaction data mapping

	function mapInteractions(rows) {
		interactionsByContactId = new Map();
		rows.forEach((row) => {
			const recordId = Number(row[QB_CONFIG.interactionFields.recordId]?.value || 0);
			const relatedContact = Number(row[QB_CONFIG.interactionFields.relatedContact]?.value || 0);
			const type = normalizeForDisplay(row[QB_CONFIG.interactionFields.type]?.value || "");
			const source = normalizeForDisplay(row[QB_CONFIG.interactionFields.source]?.value || "");
			if (!relatedContact) return;

			const existing = interactionsByContactId.get(relatedContact) || [];
			existing.push({ recordId, relatedContact, type, source });
			interactionsByContactId.set(relatedContact, existing);
		});
	}

	// Interaction rendering

	function renderInteractionOverview(filterRow) {
		if (!interactionsByContactId.size) {
			interactionOverview.style.display = "block";
			interactionOverview.innerHTML = '<div class="empty-state small"><p>No interactions found for uploaded contacts.</p></div>';
			return;
		}

		const rows = [];
		allSheetObjs.forEach((item) => {
			const contactId = Number(item.linkedRecordId);
			if (!contactId) return;
			if (typeof filterRow === "number" && filterRow !== item.rowIndex) return;

			const interactions = interactionsByContactId.get(contactId) || [];
			if (!interactions.length) return;

			interactions.forEach((oppty) => {
				rows.push(`
					<tr>
						<td>${item.rowIndex}</td>
						<td>${escapeHtml(item.name || "-")}</td>
						<td>${contactId}</td>
						<td>${oppty.recordId}</td>
						<td>${escapeHtml(oppty.type || "-")}</td>
						<td>${escapeHtml(oppty.source || "-")}</td>
					</tr>
				`);
			});
		});

		if (!rows.length) {
			interactionOverview.style.display = "block";
			interactionOverview.innerHTML = '<div class="empty-state small"><p>No interactions available for this selection.</p></div>';
			return;
		}

		interactionOverview.style.display = "block";
		interactionOverview.innerHTML = `
			<div class="overview-header">
				<h3>Existing Interactions</h3>
				<button id="hideInteractionOverviewBtn" class="btn-secondary">Hide</button>
			</div>
			<div class="table-wrap">
				<table class="overview-table">
					<thead>
						<tr>
							<th>Row</th>
							<th>Contact</th>
							<th>Contact ID</th>
							<th>Interaction ID</th>
							<th>Type</th>
							<th>Source</th>
						</tr>
					</thead>
					<tbody>
						${rows.join("")}
					</tbody>
				</table>
			</div>
		`;
	}

	// Bulk interaction actions

	async function removeAllInteractionsForUploaded() {
		showInteractionStatus("Removing interactions for uploaded contacts...", false);

		try {
			const contactIds = await ensureContactIdsForAllUploaded(false);
			const rows = await queryExistingInteractions(contactIds);
			if (!rows.length) {
				showInteractionStatus("No existing interactions to remove.", false);
				return;
			}

			const where = rows
				.map((row) => Number(row[QB_CONFIG.interactionFields.recordId]?.value || 0))
				.filter(Boolean)
				.map((id) => `{${QB_CONFIG.interactionFields.recordId}.EX.${id}}`)
				.join("OR");

			if (!where) {
				showInteractionStatus("No removable interactions found.", false);
				return;
			}

			const response = await qbRequest("https://api.quickbase.com/v1/records/delete", "POST", {
				from: QB_CONFIG.interactionTableId,
				where
			}, QB_CONFIG.interactionTableId);

			assertNoLineErrors(response, "Delete interactions");
			interactionOverview.style.display = "none";
			interactionsByContactId = new Map();
			showInteractionStatus(`Removed ${rows.length} interactions.`, false);
		} catch (error) {
			showInteractionStatus(`Failed to remove interactions: ${error.message || error}`, true);
		}
	}

	async function addInteractionToAllUploaded() {
		const interactionType = getInteractionTypeValue();
		const showName = getShowNameValue();
		const sourceDateInput = normalizeForDisplay(uploadSourceDate ? uploadSourceDate.value : "");
		const source = getInteractionSourceValue(showName, interactionSource ? interactionSource.value : "");
		let sourceDate = "";
		try {
			sourceDate = getUploadScopedSourceDateValue(sourceDateInput);
		} catch (error) {
			showInteractionStatus(error.message || error, true);
			return;
		}
		if (!interactionType) {
			showInteractionStatus("Select an interaction type first.", true);
			return;
		}

		setButtonLoading(addInteractionAllBtn, "Adding interactions...");
		showInteractionStatus("Linking contacts and creating interactions...", false);

		try {
			const contactIds = await ensureContactIdsForAllUploaded(true);
			const createdIds = await createInteractionsForContactIds(contactIds, interactionType, showName, source, sourceDate);
			addInteractionAllBtn.textContent = `Added ${createdIds.length}`;
			setTimeout(() => {
				resetButton(addInteractionAllBtn);
			}, 1200);
			showInteractionStatus(
				`Created ${createdIds.length} interactions with show name \"${showName || "(blank)"}\", type \"${interactionType}\", source \"${source || "(blank)"}\" and source date \"${sourceDate || "(blank)"}\".`,
				false
			);
			renderResults();
		} catch (error) {
			resetButton(addInteractionAllBtn);
			addInteractionAllBtn.textContent = "Retry Bulk Add";
			showInteractionStatus(`Bulk interaction create failed: ${error.message || error}`, true);
		}
	}

	async function viewExistingInteractionsForUploaded(filterRow) {
		setButtonLoading(viewInteractionAllBtn, "Loading interactions...");
		showInteractionStatus("Loading existing interactions...", false);

		try {
			const contactIds = await ensureContactIdsForAllUploaded(false);
			const rows = await queryExistingInteractions(contactIds);
			mapInteractions(rows);
			renderInteractionOverview(typeof filterRow === "number" ? filterRow : undefined);
			showInteractionStatus(`Loaded ${rows.length} interactions.`, false);
			resetButton(viewInteractionAllBtn);
		} catch (error) {
			resetButton(viewInteractionAllBtn);
			showInteractionStatus(`Could not load interactions: ${error.message || error}`, true);
		}
	}

	// Single contact interaction

	async function addInteractionForSingleContact(row, btn) {
		setButtonLoading(btn, "Preparing...");

		try {
			const recordId = await ensureLinkedContactRecordForRow(Number(row));
			resetButton(btn);
			showContactStatus(`Ready to add interaction for contact ${recordId}.`, false);
			openInteractionModalForRow(row);
		} catch (error) {
			console.error("[Interaction Modal Prepare] Failed", {
				row,
				error: getAjaxErrorMessage(error)
			});
			resetButton(btn);
			btn.textContent = "Retry Add Interaction";
			showInteractionStatus(`Could not prepare interaction form: ${getAjaxErrorMessage(error)}`, true);
		}
	}

	async function submitInteractionFromModal() {
		if (!interactionModalState || interactionModalState.mode !== "single") return;

		const row = Number(interactionModalState.row);
		const item = getObjectByRow(row);
		if (!item) return;

		const showName = normalizeForDisplay(interactionFieldShowName.value);
		const interactionType = normalizeForDisplay(interactionFieldType.value);
		const sourceDateInput = normalizeForDisplay(uploadSourceDate ? uploadSourceDate.value : "");
		let sourceDate = "";
		try {
			sourceDate = getUploadScopedSourceDateValue(sourceDateInput);
		} catch (error) {
			showInteractionStatus(error.message || error, true);
			return;
		}
		const selectedSource = getInteractionSourceValue(showName, interactionFieldSource.value);
		const selectedRelatedContactId = Number(interactionFieldContactId.dataset.contactId || 0);

		if (!interactionType) {
			showInteractionStatus("Select an interaction type first.", true);
			return;
		}

		setButtonLoading(interactionModalSubmitBtn, "Creating...");

		try {
			const ids = selectedRelatedContactId > 0
				? [selectedRelatedContactId]
				: item.linkedRecordId
					? [item.linkedRecordId]
					: await createContactsForItems([item]);

			const contactId = Number(selectedRelatedContactId || item.linkedRecordId || ids[0]);
			if (!contactId) {
				throw new Error("Unable to link this contact to QuickBase.");
			}

			const payload = {
				to: QB_CONFIG.interactionTableId,
				data: [{
					[QB_CONFIG.interactionFields.showName]: { value: showName },
					[QB_CONFIG.interactionFields.relatedContact]: { value: contactId },
					[QB_CONFIG.interactionFields.type]: { value: interactionType },
					[QB_CONFIG.interactionFields.source]: { value: selectedSource },
					[QB_CONFIG.interactionFields.sourceDate]: { value: sourceDate }
				}]
			};

			const response = await qbRequest("https://api.quickbase.com/v1/records", "POST", payload, QB_CONFIG.interactionTableId);
			assertNoLineErrors(response, "Create interaction");
			const createdIds = (response && response.metadata && response.metadata.createdRecordIds) || [];
			showInteractionStatus(
				`Added ${createdIds.length} interaction for row ${item.rowIndex}.`,
				false
			);
			showContactStatus(`Interaction created for row ${item.rowIndex}.`, false);
			markButtonSuccess(interactionModalSubmitBtn, "Created");
			setTimeout(closeInteractionModal, SUCCESS_RESET_MS);
			renderResults();
		} catch (error) {
			console.error("[Single Interaction Create] Failed", {
				row,
				showName,
				relatedContactId: Number(interactionFieldContactId.dataset.contactId || 0),
				relatedContactName: interactionFieldContactId.value,
				type: interactionType,
				source: selectedSource,
				sourceDate,
				error: getAjaxErrorMessage(error)
			});
			resetButton(interactionModalSubmitBtn);
			interactionModalSubmitBtn.textContent = "Retry Create";
			showInteractionStatus(`Failed to add interaction: ${getAjaxErrorMessage(error)}`, true);
		}
	}

	// Contact card editing

	function saveContactEdit(card) {
		const row = Number(card.dataset.row);
		const item = getObjectByRow(row);
		if (!item) return;

		const firstNameInput = card.querySelector(".edit-first-name");
		const lastNameInput = card.querySelector(".edit-last-name");
		const emailInput = card.querySelector(".edit-email");

		const firstName = normalizeForDisplay(firstNameInput.value);
		const lastName = normalizeForDisplay(lastNameInput.value);
		const email = normalizeForDisplay(emailInput.value);
		const fullName = [firstName, lastName].filter(Boolean).join(" ");

		if (!fullName && !email) {
			showContactStatus(`Row ${row} cannot be empty.`, true);
			return;
		}

		item.firstName = firstName;
		item.lastName = lastName;
		item.name = fullName;
		item.email = email;
		item.nameNorm = normalizeForMatch(fullName);
		item.emailNorm = normalizeForMatch(email);
		item.affiliation = item.affiliation || "";
		item.affiliationNorm = normalizeForMatch(item.affiliation);
		item.combinedNorm = normalizeForMatch(`${fullName} ${email} ${item.affiliation}`);

		const match = getMatchByRow(row);
		if (match) {
			match.item = item;
			matchByRow.set(Number(row), match);
		}

		renderResults();
		showContactStatus(`Saved edits for row ${row}.`, false);
	}

	function toggleEditCard(card, shouldShow) {
		const details = card ? card.querySelector(".contact-card-shell") : null;
		if (details && shouldShow) {
			details.open = true;
		}
		card.classList.toggle("editing", shouldShow);
	}

	// Contact card rendering

	function renderContactCard(item, opts) {
		const isDuplicate = !!opts.isDuplicate;
		const match = opts.match || null;
		const confidenceScore = Number(isDuplicate ? (match && match.score) : item.matchConfidence) || 0;
		const contactId = Number(item.linkedRecordId || (match && match.by && match.by.recordId) || 0);
		const contactUrl = qbContactRecordUrl(contactId);
		const contactLink = contactUrl
			? `<a class="contact-link-btn" href="${contactUrl}" target="_blank" rel="noopener noreferrer">Open QuickBase Record</a>`
			: "";
		const scoreBadgeClass = confidenceScore >= 95 ? "badge-score-high" : confidenceScore >= 90 ? "badge-score-medium" : "badge-score-low";
		const scoreBadge = isDuplicate
			? `<span class="contact-badge badge-score ${scoreBadgeClass}">${confidenceScore}%</span>`
			: `<span class="contact-badge badge-score ${scoreBadgeClass}">${confidenceScore}%</span>`;

		const matchDetails = isDuplicate
			? `
				<div class="match-details">
					<strong>Matched to:</strong>
					<span class="qb-link">${escapeHtml(match.by.name)} (${escapeHtml(match.by.email || "no email")})</span>
					<span class="qb-link">Affiliation: ${escapeHtml(match.by.affiliation || "None")}</span>
					${match.matchType === "exact-email"
						? '<span class="match-type-label">Exact Email</span>'
						: match.matchType === "exact-name-affiliation"
							? '<span class="match-type-label">Exact Name + Affiliation</span>'
							: '<span class="match-type-label">Fuzzy Match</span>'}
				</div>
			`
			: "";

		const mainAction = isDuplicate
			? `
				<button class="btn-secondary btn-update" data-action="update" data-row="${item.rowIndex}" data-record-id="${contactId}">
					Update Existing
				</button>
				<button class="btn-secondary" data-action="create-anyway" data-row="${item.rowIndex}">
					Create Anyway
				</button>
			`
			: `
				<button class="btn-secondary btn-create" data-action="create" data-row="${item.rowIndex}">
					Create Contact
				</button>
			`;

		return `
			<div class="contact-card" data-row="${item.rowIndex}">
				<details class="contact-card-shell" ${isDuplicate ? "open" : ""}>
					<summary class="contact-summary">
						<div class="contact-summary-main">
							<h3 class="contact-name">${escapeHtml(item.name || "Unnamed Contact")}</h3>
							<p class="contact-email">${escapeHtml(item.email || "No email")}</p>
							<p class="contact-meta">Affiliation: ${escapeHtml(item.affiliation || "None")}</p>
							${contactId ? `<p class="contact-meta">Linked Contact ID: ${contactId}</p>` : '<p class="contact-meta">Not linked to QuickBase yet</p>'}
						</div>
						<div class="contact-summary-side">${scoreBadge}</div>
					</summary>
					<div class="contact-body">
						${matchDetails}
						<div class="contact-link-wrap">${contactLink}</div>
						<div class="edit-panel">
							<div class="edit-grid">
								<div class="form-group">
									<label>First Name</label>
									<input class="edit-first-name" type="text" value="${escapeHtml(item.firstName || "")}" />
								</div>
								<div class="form-group">
									<label>Last Name</label>
									<input class="edit-last-name" type="text" value="${escapeHtml(item.lastName || "")}" />
								</div>
								<div class="form-group form-group-wide">
									<label>Email</label>
									<input class="edit-email" type="email" value="${escapeHtml(item.email || "")}" />
								</div>
							</div>
							<div class="edit-actions">
								<button class="btn-secondary" data-action="save-edit" data-row="${item.rowIndex}">Save Contact</button>
								<button class="btn-secondary" data-action="cancel-edit" data-row="${item.rowIndex}">Cancel</button>
							</div>
						</div>
						<div class="contact-actions">
							<span class="contact-row-info">Row ${item.rowIndex}</span>
							<button class="btn-secondary" data-action="start-edit" data-row="${item.rowIndex}">Edit</button>
							<button class="btn-secondary" data-action="add-interaction" data-row="${item.rowIndex}">Add Interaction</button>
							${mainAction}
						</div>
					</div>
				</details>
			</div>
		`;
	}

	// Results rendering

	function renderDuplicates() {
		if (currentMatches.length === 0) {
			duplicatesList.innerHTML = `
				<div class="empty-state">
					<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
						<path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/>
					</svg>
					<p>No duplicate contacts found.</p>
				</div>
			`;
			return;
		}

		duplicatesList.innerHTML = currentMatches.map((m) => renderContactCard(m.item, { isDuplicate: true, match: m })).join("");
	}

	function renderNewContacts() {
		if (nonMatches.length === 0) {
			newContactsList.innerHTML = `
				<div class="empty-state">
					<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
						<path d="M20 13V6a2 2 0 00-2-2H6a2 2 0 00-2 2v7m16 0v5a2 2 0 01-2 2H6a2 2 0 01-2-2v-5"/>
					</svg>
					<p>No new contacts to add.</p>
				</div>
			`;
			return;
		}

		newContactsList.innerHTML = nonMatches.map((item) => renderContactCard(item, { isDuplicate: false })).join("");
	}

	function renderResults() {
		rebuildSheetObjectLookup();
		rebuildMatchLookup();

		const matchedRows = new Set(currentMatches.map((m) => Number(m.item.rowIndex)));
		nonMatches = allSheetObjs.filter((obj) => !matchedRows.has(Number(obj.rowIndex)));
		duplicateCount.textContent = String(currentMatches.length);
		newContactCount.textContent = String(nonMatches.length);
		totalCount.textContent = String(allSheetObjs.length);

		renderDuplicates();
		renderNewContacts();

		updateAllBtn.style.display = currentMatches.length ? "block" : "none";
		addAllBtn.style.display = nonMatches.length ? "block" : "none";
		resultsContainer.style.display = "block";
	}

	// Export

	function getExportTimestamp() {
		return CFUtils.getExportTimestamp();
	}

	function createContactExportRows(items) {
		return CFUtils.createContactExportRows(items);
	}

	function exportMatchResultsToExcel() {
		if (!allSheetObjs.length) {
			showContactStatus("Nothing to export yet. Process a file first.", true);
			return;
		}

		const duplicateRows = createContactExportRows(currentMatches.map((m) => m.item));
		const newRows = createContactExportRows(nonMatches);

		const workbook = XLSX.utils.book_new();
		const duplicateSheet = duplicateRows.length
			? XLSX.utils.json_to_sheet(duplicateRows)
			: XLSX.utils.aoa_to_sheet([["No duplicate contacts in this run"]]);
		const newContactsSheet = newRows.length
			? XLSX.utils.json_to_sheet(newRows)
			: XLSX.utils.aoa_to_sheet([["No new contacts in this run"]]);

		XLSX.utils.book_append_sheet(workbook, duplicateSheet, "Duplicates");
		XLSX.utils.book_append_sheet(workbook, newContactsSheet, "New Contacts");

		const fileName = `contact-match-export-${getExportTimestamp()}.xlsx`;
		XLSX.writeFile(workbook, fileName);
		showContactStatus(`Exported ${duplicateRows.length} duplicates and ${newRows.length} new contacts to ${fileName}.`, false);
	}

	// Card action dispatcher

	function handleCardActions(e) {
		const statCard = e.target.closest(".stat-clickable");
		if (statCard) {
			const targetId = statCard.getAttribute("data-scroll-to");
			const targetSection = document.getElementById(targetId);
			if (targetSection) {
				targetSection.scrollIntoView({ behavior: "smooth", block: "start" });
			}
			return;
		}

		const btn = e.target.closest("[data-action]");
		if (!btn) return;

		const action = btn.getAttribute("data-action");
		const row = btn.getAttribute("data-row");
		const recordId = btn.getAttribute("data-record-id");
		const card = btn.closest(".contact-card");

		switch (action) {
			case "update":
				updateContact(row, recordId, btn, card);
				break;
			case "create":
			case "create-anyway":
				createContact(row, btn);
				break;
			case "start-edit":
				toggleEditCard(card, true);
				break;
			case "cancel-edit":
				toggleEditCard(card, false);
				break;
			case "save-edit":
				saveContactEdit(card);
				break;
			case "add-interaction":
				addInteractionForSingleContact(row, btn);
				break;
		}
	}

	// Event wiring

	if (loadDemoBtn) {
		loadDemoBtn.addEventListener("click", loadDemoUpload);
	}

	processBtn.addEventListener("click", async function () {
		if (storedContacts.length === 0) {
			showContactStatus("Contacts are still loading from QuickBase. Please wait.", true);
			return;
		}

		const file = fileInput.files && fileInput.files[0];
		if (!file) {
			showContactStatus("Please choose an Excel file first.", true);
			return;
		}

		showContactStatus("Processing file...", false);

		try {
			// Incoming contacts from the user Excel file
			const rows = await parseWorkbook(file);
			allSheetObjs = sheetRowsToObjects(rows);
			rebuildSheetObjectLookup();
			if (!allSheetObjs.length) {
				showContactStatus("No valid rows found in the file.", true);
				return;
			}

			uploadScopedSourceDate = "";
			currentMatches = searchMatches(allSheetObjs);
			interactionsByContactId = new Map();
			interactionOverview.style.display = "none";
			interactionOverview.innerHTML = "";
			renderResults();
			showContactStatus(`Processed ${allSheetObjs.length} uploaded contacts.`, false);
			showInteractionStatus("Interaction tools are ready.", false);
		} catch (error) {
			showContactStatus("Error reading file.", true);
		}
	});

	updateAllBtn.addEventListener("click", updateAllContacts);
	addAllBtn.addEventListener("click", createAllContacts);
	if (exportResultsBtn) {
		exportResultsBtn.addEventListener("click", exportMatchResultsToExcel);
	}
	addInteractionAllBtn.addEventListener("click", addInteractionToAllUploaded);
	removeInteractionAllBtn.addEventListener("click", removeAllInteractionsForUploaded);
	if (viewInteractionAllBtn) {
		viewInteractionAllBtn.addEventListener("click", function () {
			viewExistingInteractionsForUploaded();
		});
	}

	if (interactionModalCloseBtn) {
		interactionModalCloseBtn.addEventListener("click", closeInteractionModal);
	}
	if (interactionModalCancelBtn) {
		interactionModalCancelBtn.addEventListener("click", closeInteractionModal);
	}
	if (interactionModalOverlay) {
		interactionModalOverlay.addEventListener("click", closeInteractionModal);
	}
	if (interactionForm) {
		interactionForm.addEventListener("submit", function (e) {
			e.preventDefault();
			submitInteractionFromModal();
		});
	}
	if (interactionShowName) {
		interactionShowName.addEventListener("input", function () {
			syncSourceControlsFromShowName();
		});
	}

	document.addEventListener("click", function (e) {
		if (e.target && e.target.id === "hideInteractionOverviewBtn") {
			interactionOverview.style.display = "none";
			return;
		}
		handleCardActions(e);
	});

	fileInput.addEventListener("change", function (e) {
		const fileName = e.target.files[0] && e.target.files[0].name;
		if (!fileName) return;
		const label = document.querySelector(".file-label");
		if (!label) return;
		label.innerHTML = `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
			<path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/>
		</svg> ${escapeHtml(fileName)}`;
	});

	window.contactFiltering = {
		getStoredContacts: function () {
			return storedContacts;
		},
		reloadContacts: loadContactsFromQB
	};

	// Initialization

	if (USE_DEMO_CONTACTS) {
		Demo.init();
	}

	showContactStatus("Initializing...", false);
	if (interactionModal) {
		interactionModal.hidden = true;
	}
	if (uploadSourceDate && !uploadSourceDate.value) {
		uploadSourceDate.value = getTodayIsoDate();
	}
	syncSourceControlsFromShowName();
	(async function initContacts() {
		try {
			await loadContactsFromQB();
		} catch (error) {
			showContactStatus(`Failed to load contacts: ${getAjaxErrorMessage(error)}`, true);
		}
	})();
})();
