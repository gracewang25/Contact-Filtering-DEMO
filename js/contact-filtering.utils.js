// contact-filtering.utils.js
// Formatting, parsing, and error helpers used across the app.

(function () {
	const ContactFilteringUtils = {
		// String normalization
		normalizeForMatch(v) {
			return String(v || "")
				.trim()
				.toLowerCase()
				.replace(/\s+/g, " ");
		},

		normalizeForDisplay(v) {
			return String(v || "").trim().replace(/\s+/g, " ");
		},

		splitName(fullName) {
			const trimmed = ContactFilteringUtils.normalizeForDisplay(fullName);
			if (!trimmed) return { firstName: "", lastName: "" };

			const parts = trimmed.split(/\s+/);
			if (parts.length === 1) return { firstName: parts[0], lastName: "" };
			return { firstName: parts[0], lastName: parts.slice(1).join(" ") };
		},

		escapeHtml(s) {
			return String(s || "").replace(/[&<>\"']/g, function (c) {
				return {
					"&": "&amp;",
					"<": "&lt;",
					">": "&gt;",
					'"': "&quot;",
					"'": "&#39;"
				}[c];
			});
		},
		// Debug, request, and error helpers
		getHeadersForLog(headers) {
			const cloned = Object.assign({}, headers || {});
			if (cloned.Authorization) {
				cloned.Authorization = `${String(cloned.Authorization).slice(0, 22)}...`;
			}
			return cloned;
		},

		getAjaxErrorMessage(err) {
			if (!err) return "Unknown QuickBase error";
			if (err.responseJSON && err.responseJSON.message) return err.responseJSON.message;
			if (err.responseJSON && err.responseJSON.description) return err.responseJSON.description;
			if (err.responseText) return err.responseText;
			if (err.statusText) return err.statusText;
			return String(err);
		},

		assertNoLineErrors(response, actionLabel) {
			const lineErrors = response && response.metadata && response.metadata.lineErrors;
			if (lineErrors && Object.keys(lineErrors).length > 0) {
				const firstLine = Object.keys(lineErrors)[0];
				const firstMsg = lineErrors[firstLine] && lineErrors[firstLine][0] ? lineErrors[firstLine][0] : "Unknown line error";
				throw new Error(`${actionLabel} failed on row ${firstLine}: ${firstMsg}`);
			}
		},

		resolveTokenExpiryMs(data) {
			const now = Date.now();
			const rawExpiry = data && (data.expiresAt || data.expiration || data.temporaryAuthorizationExpiration);

			if (rawExpiry) {
				const parsed = Date.parse(rawExpiry);
				if (!Number.isNaN(parsed)) {
					return parsed - 15000;
				}
			}

			const expiresInSeconds = Number(data && data.expiresIn);
			if (!Number.isNaN(expiresInSeconds) && expiresInSeconds > 0) {
				return now + (expiresInSeconds * 1000) - 15000;
			}

			return now + (5 * 60 * 1000) - 15000;
		},

		// Generic utils
		chunkArray(items, chunkSize) {
			const out = [];
			for (let i = 0; i < items.length; i += chunkSize) {
				out.push(items.slice(i, i + chunkSize));
			}
			return out;
		},

		toConfidenceFromScore(fuseScore) {
			const confidence = Math.round((1 - (fuseScore || 0)) * 100);
			return Math.max(0, Math.min(100, confidence));
		},

		// Date and export helpers
		getTodayIsoDate() {
			return new Date().toISOString().slice(0, 10);
		},

		normalizeDateForQuickBase(rawDate) {
			const value = ContactFilteringUtils.normalizeForDisplay(rawDate);
			if (!value) return ContactFilteringUtils.getTodayIsoDate();

			if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
				return value;
			}

			const parsed = new Date(value);
			if (!Number.isNaN(parsed.getTime())) {
				return parsed.toISOString().slice(0, 10);
			}

			throw new Error("Interaction Date must be a valid date.");
		},

		getExportTimestamp() {
			const now = new Date();
			const pad = (n) => String(n).padStart(2, "0");
			return `${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}`;
		},

		createContactExportRows(items) {
			return items.map((item) => ({
				Name: ContactFilteringUtils.normalizeForDisplay(item && item.name),
				Email: ContactFilteringUtils.normalizeForDisplay(item && item.email),
				Affiliation: ContactFilteringUtils.normalizeForDisplay(item && item.affiliation)
			}));
		}
	};

	window.ContactFilteringUtils = ContactFilteringUtils;
})();
