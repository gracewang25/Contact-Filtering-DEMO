// contact-filtering.data.js
// Workbook parsing and duplicate detection logic.

(function () {
	const ContactFilteringData = {
		// Workbook parsing
		parseWorkbook(file) {
			return new Promise((resolve, reject) => {
				const reader = new FileReader();
				reader.onload = function (e) {
					try {
						const data = e.target.result;
						const workbook = XLSX.read(data, { type: "binary" });
						const firstSheetName = workbook.SheetNames[0];
						const worksheet = workbook.Sheets[firstSheetName];
						const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
						resolve(rows);
					} catch (error) {
						reject(error);
					}
				};
				reader.onerror = function (err) {
					reject(err);
				};
				reader.readAsBinaryString(file);
			});
		},

		// Row transformation
		sheetRowsToObjects(rows) {
			const utils = window.ContactFilteringUtils;
			const objs = [];
			let startIndex = 0;

			// Default cols: full name, email, affiliation
			let nameCol = 0;
			let emailCol = 1;
			let affiliationCol = 2;
 
			if (rows.length > 0) {
				const first = rows[0].map((c) => String(c || "").toLowerCase().trim());
				const looksLikeHeader = first.some((c) => c.includes("name") || c.includes("email") || c.includes("affiliation"));

				if (looksLikeHeader) {
					startIndex = 1;

					const nameIndex = first.findIndex((c) => c.includes("full name") || c === "name" || c.includes("contact name"));
					const emailIndex = first.findIndex((c) => c.includes("email address") || c === "email" || c.includes("e-mail"));
					const affiliationIndex = first.findIndex((c) => c.includes("affiliation"));

					if (nameIndex >= 0) nameCol = nameIndex;
					if (emailIndex >= 0) emailCol = emailIndex;
					if (affiliationIndex >= 0) affiliationCol = affiliationIndex;
				}
			}

			for (let i = startIndex; i < rows.length; i += 1) {
				const r = rows[i];
				if (!r) continue;

				const nameRaw = utils.normalizeForDisplay(r[nameCol]);
				const email = utils.normalizeForDisplay(r[emailCol]);
				const affiliation = utils.normalizeForDisplay(r[affiliationCol]);
				if (!nameRaw && !email && !affiliation) continue;

				const names = utils.splitName(nameRaw);
				const displayName = [names.firstName, names.lastName].filter(Boolean).join(" ") || nameRaw;

				objs.push({
					rowIndex: i + 1,
					firstName: names.firstName,
					lastName: names.lastName,
					name: displayName,
					email: email,
					affiliation,
					nameNorm: utils.normalizeForMatch(displayName),
					emailNorm: utils.normalizeForMatch(email),
					affiliationNorm: utils.normalizeForMatch(affiliation),
					combinedNorm: utils.normalizeForMatch(`${displayName} ${email} ${affiliation}`),
					linkedRecordId: null,
					raw: r
				});
			}
			return objs;
		},

		// Duplicate detection
		searchMatches(sheetObjs, storedContacts, duplicateScoreThreshold) {
			const utils = window.ContactFilteringUtils;
			const matches = [];
			const usedRowIndexes = new Set();
			const emailIndex = new Map();
			const nameAffiliationIndex = new Map();

			storedContacts.forEach((contact) => {
				if (contact.emailNorm && !emailIndex.has(contact.emailNorm)) {
					emailIndex.set(contact.emailNorm, contact);
				}
				if (contact.nameNorm && contact.affiliationNorm) {
					const key = `${contact.nameNorm}|${contact.affiliationNorm}`;
					if (!nameAffiliationIndex.has(key)) {
						nameAffiliationIndex.set(key, contact);
					}
				}
			});

			const fuse = new Fuse(storedContacts, {
				keys: ["nameNorm", "emailNorm", "affiliationNorm", "combinedNorm"],
				includeScore: true,
				threshold: 0.40,
				ignoreLocation: true,
				minMatchCharLength: 2
			});

			sheetObjs.forEach((item) => {
				let best = null;
				item.matchConfidence = 0;

				if (item.emailNorm) {
					const exactEmail = emailIndex.get(item.emailNorm);
					if (exactEmail) {
						item.matchConfidence = 100;
						best = {
							item,
							by: exactEmail,
							score: 100,
							matchType: "exact-email"
						};
					}
				}

				if (!best) {
					const exactNameAffiliation = (item.nameNorm && item.affiliationNorm)
						? nameAffiliationIndex.get(`${item.nameNorm}|${item.affiliationNorm}`) 
						: null;
					if (exactNameAffiliation) {
						item.matchConfidence = 98;
						best = {
							item,
							by: exactNameAffiliation,
							score: 98,
							matchType: "exact-name-affiliation"
						};
					}
				}

				if (!best) {
					const query = [item.nameNorm, item.emailNorm, item.affiliationNorm].filter(Boolean).join(" ");
					if (!query) return;
					const results = fuse.search(query, { limit: 1 });
					if (results.length > 0) {
						const top = results[0];
						const confidence = utils.toConfidenceFromScore(top.score);
						item.matchConfidence = confidence;
						if (confidence >= duplicateScoreThreshold) {
							best = {
								item,
								by: top.item,
								score: confidence,
								matchType: "fuzzy"
							};
						}
					}
				}

				if (best) {
					item.linkedRecordId = best.by.recordId;
					matches.push(best);
					usedRowIndexes.add(item.rowIndex);
				}
			});

			const sortedMatches = matches.sort((a, b) => b.score - a.score);
			const unmatchedItems = sheetObjs.filter((obj) => !usedRowIndexes.has(obj.rowIndex));

			return {
				matches: sortedMatches,
				nonMatches: unmatchedItems,
				matchByRow: new Map(sortedMatches.map((match) => [Number(match.item.rowIndex), match]))
			};
		}
	};

	window.ContactFilteringData = ContactFilteringData;
})();
