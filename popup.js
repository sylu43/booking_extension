document.getElementById("run").addEventListener("click", () => {
	chrome.tabs.query({active: true}, ([tab]) => {
		chrome.tabs.sendMessage(tab.id, { action: "run" });
	});
});
