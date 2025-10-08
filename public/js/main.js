// Maneja el click en "Exportar Excel" para usar el layout nuevo /export/tpa-form
(function(){
	function getSampleIdFromPage(){
		const input = document.querySelector('input[name="sample_id"]');
		if (input && input.value) return input.value.trim();
		return '';
	}
	function handleExportClick(ev){
		ev.preventDefault();
		let sampleId = getSampleIdFromPage();
		if (!sampleId){
			sampleId = window.prompt('Ingrese sample_id para exportar:','1') || '';
			sampleId = sampleId.trim();
		}
		if (!sampleId) return;
		window.location.href = `/export/tpa-form?sample_id=${encodeURIComponent(sampleId)}`;
	}
	document.addEventListener('click', function(e){
		const a = e.target.closest('[data-export-tpa]');
		if (a) handleExportClick(e);
	});
})();
