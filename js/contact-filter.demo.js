// contact-filter.demo.js
// Enables demo mode — loads sample data and mocks QB API calls so the full UI works without a live connection.
// To switch to live QuickBase: remove this file's <script> tag in the HTML. Nothing else changes.

(function () {
	let nextRecordId = 9100;

	// Existing contacts that the app "loads from QuickBase" on startup
	const EXISTING_CONTACTS = [
		{ recordId: 9001, firstName: "Avery",      lastName: "Nguyen",  email: "avery.nguyen@northlakebio.com",   affiliation: "Northlake Bio" },
		{ recordId: 9002, firstName: "Jordan",     lastName: "Patel",   email: "jordan.patel@quantumlabs.io",     affiliation: "Quantum Labs" },
		{ recordId: 9003, firstName: "Sam",        lastName: "Rivera",  email: "sam.rivera@helixpoint.ai",        affiliation: "HelixPoint AI" },
		{ recordId: 9004, firstName: "Taylor",     lastName: "Kim",     email: "taylor.kim@veridianresearch.org", affiliation: "Veridian Research" },
		{ recordId: 9005, firstName: "Morgan",     lastName: "Lee",     email: "morgan.lee@arcadiabio.net",       affiliation: "Arcadia Bio" },
		{ recordId: 9006, firstName: "Tiffany",    lastName: "Snith",   email: "tiffany.smith@thermofisher.com",  affiliation: "Thermo Fisher Scientific" },
		{ recordId: 9007, firstName: "Anne Marie", lastName: "Oniel",   email: "anne-marie.oneil@coniferlsg.org", affiliation: "Conifer Life Sciences Group" }
	];

	// Partially overlaps EXISTING_CONTACTS to showcase duplicate detection on upload
	const UPLOAD_CONTACTS = [
		{ firstName: "Avery",   lastName: "Nguyen", email: "avery.nguyen@northlakebio.com",  affiliation: "Northlake Bio" },
		{ firstName: "Tiffany", lastName: "Smith",  email: "tiffany.smith@thermofisher.com", affiliation: "Thermo Fisher Scientific" },
		{ firstName: "Jordan",  lastName: "Patel",  email: "j.patel@quantumlabs.io",         affiliation: "Quantum Labs" },
		{ firstName: "Samuel",  lastName: "Rivera", email: "s.rivera@helixpoint.ai",         affiliation: "HelixPoint AI" },
		{ firstName: "Priya",   lastName: "Mehta",  email: "priya.mehta@synbioworks.com",    affiliation: "SynBio Works" },
		{ firstName: "Daniel",  lastName: "Osei",   email: "daniel.osei@genomiqs.io",        affiliation: "Genomiqs" },
		{ firstName: "Claire",  lastName: "Huang",  email: "claire.huang@vantabio.com",      affiliation: "Vanta Bio" }
	];

	// Minimal QB mock — returns realistic metadata so all UI flows (create, update, delete) complete without errors
	function qbRequest(url, method, payload, config) {
		return new Promise(function (resolve) {
			setTimeout(function () {
				if (url.includes("/records/delete")) {
					resolve({});
					return;
				}

				if (url.includes("/records/query")) {
					resolve({ data: [] });
					return;
				}

				if (payload && payload.data && payload.data.length) {
					const recordIdField = config && config.contactFields && config.contactFields.recordId;
					const isUpdate = recordIdField && payload.data.some(function (r) {
						return r[recordIdField] && Number(r[recordIdField].value);
					});

					if (isUpdate) {
						const updatedIds = payload.data
							.map(function (r) { return Number((r[recordIdField] || {}).value || 0); })
							.filter(Boolean);
						resolve({ metadata: { createdRecordIds: [], updatedRecordIds: updatedIds, unchangedRecordIds: [] } });
						return;
					}

					const createdIds = payload.data.map(function () { return nextRecordId++; });
					resolve({ metadata: { createdRecordIds: createdIds, updatedRecordIds: [], unchangedRecordIds: [] } });
					return;
				}

				resolve({});
			}, 400);
		});
	}

	// Runs once on init — reveals demo UI elements and pre-fills the show name field
	function init() {
		const demoBannerEl = document.getElementById("demoBanner");
		if (demoBannerEl) demoBannerEl.hidden = false;

		const loadDemoBtnEl = document.getElementById("loadDemoBtn");
		if (loadDemoBtnEl) loadDemoBtnEl.hidden = false;

		const interactionShowName = document.getElementById("interactionShowName");
		if (interactionShowName && !interactionShowName.value) {
			interactionShowName.value = "BIO International 2026";
		}
	}

	window.ContactFilteringDemo = {
		useDemo: true,
		existingContacts: EXISTING_CONTACTS,
		uploadContacts: UPLOAD_CONTACTS,
		qbRequest: qbRequest,
		init: init
	};
})();
