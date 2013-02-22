<?php
	header('Content-Type: text/html; charset=utf-8');
?>
	<div>
		<h1 style="display: inline">:: Excel importing ::</h1>
		&emsp;
		<div id="loading" style="display: inline"></div>
	</div><div style="clear: both"></div>
	<hr/>
	<div id="loader"></div>

<script src="http://code.jquery.com/jquery-latest.js"></script>
<script type="text/javascript">

	document.title = "Excel import to MySQL";

	$(function() {
		var process = $.ajax({
			url:	'http://localhost/excel_import.php',
			dataType:	'html',
			beforeSend:	function() {
				$("#loading").html("<img src=ajax-loader.gif>");
			},
			success: function(data) {
				$("#loading").empty();
				$("#loader").html(data);
			}
		});

		process.done( function() {
//			$('#loader').append('<p>Process done!!</p>');
		});

		process.fail( function(status, message) {
			$('#loader').append('<p>Process failed!!</p><p>Status : ' + status + '</p><p>Message : ' + message + '</p>');
		});

		process.always( function() {
//			$('#loader').append('<p>Process finished!!</p>');
		});
	});
</script>