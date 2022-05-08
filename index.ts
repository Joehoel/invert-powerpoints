document.querySelector<HTMLButtonElement>("#open")?.addEventListener("click", async () => {
    const handle = await window.showDirectoryPicker();
    const values = handle.values();
});
