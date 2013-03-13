<?php
	include('PHPExcel.php');
	include('PHPExcel/IOFactory.php');
	include('PHPExcel/Reader/Excel2007.php');
	include('PHPExcel/Worksheet.php');
	include('ChunkReadFilter.php');

	if ( !function_exists( 'debug_log' ) ) {
		function debug_log( $msg, $file = './_debug/debug.txt' ) {
			$msg = gmdate( 'Y-m-d H:i:s' ) . ' ' . print_r( $msg, TRUE ) . "\n";
			error_log( $msg, 3, $file );
		}
	}

	function getFloor($data = array()) {
		list($junk, $floor) = explode('ชั้น', $data);

		if(!$floor) return '';
		else		return $floor;
	}

	function makeTime($time = 0) {
		if(!$time)
			return 0;

		list($d, $m, $y) = explode('/', $time);
		$intTime	= mktime(0,0,0, $m, $d, $y);

		return date('Y-m-d', $intTime);
	}

	gc_disable();
	ini_set('memory_limit', '-1');
	set_time_limit(0);

	$cacheMethod = PHPExcel_CachedObjectStorageFactory::cache_to_phpTemp;
	$cacheSettings = array( 'memoryCacheSize'  => '16MB' );

	PHPExcel_Settings::setCacheStorageMethod($cacheMethod, $cacheSettings);

	// benchmark
	$start = microtime(true);

	$filename		= realpath('./excels/sample-latest.xls');
	$name			= pathinfo($filename, PATHINFO_BASENAME);
	$ext			= pathinfo($filename, PATHINFO_EXTENSION);

	$reader			= PHPExcel_IOFactory::createReaderForFile($filename);
	$canRead		= $reader->canRead($filename);

	if(!$canRead) return false;

	// To read only data, not format and formula
	$reader->setReadDataOnly(true);
	$obj		= $reader->load($filename);

	// To read by chunk
//	$setFilter		= chunkReadFilter::setFilter($reader);
//	$chunkSize		= 1000;

	/* ---------------------------- Output Zone ---------------------------- */

	/*for($readRow = 2; $readRow <= 100 ; $readRow += $chunkSize) {
		$chunkFilter->setRows($readRow, $chunkSize);
		$excel		= $reader->load($filename);
		$worksheet	= $excel->setActiveSheetIndex(0);
		$rows		= $worksheet->getRowIterator();
		foreach($rows as $row) {
			$cellIterator	= $row->getCellIterator();
			$cellIterator->setIterateOnlyExistingCells(false);
			foreach($cellIterator as $cell) {
				if(!is_null($cell)) {
					echo '<td>'.$cell->getValue().'</td>';
				}
			}
		}
	}*/

	// TODO check database connection every time
	//<editor-fold desc="Database connection">
	$host		= "localhost";
	$username	= "root";
	$password	= "1234";
	$db_name	= 'work1';
	$prefix		= 'jos' . '_'; // i.e. jos_
	$mysqli 	= new mysqli($host, $username, $password, $db_name);
	//</editor-fold>

	if($mysqli->connect_errno) {
		echo "<h3>Cannot connect to database<br/>[ <font color=red>".$mysqli->connect_errno."</font> - ".$mysqli->connect_error." ]!</h3><hr/>";
	}
	else {
		echo "<h3 style='color: #32cd32'>Connected!</h3><hr/>";
		echo "<h3>Select database : ".$db_name."</h3>";
		$mysqli->set_charset('utf8');

		$worksheet		= $obj->setActiveSheetIndex(0);
		$rowLast		= $worksheet->getHighestRow();
		$colLast		= $worksheet->getHighestColumn();
		$rows			= $worksheet->rangeToArray('A2:'.$colLast.$rowLast,null,true,true,true);

		//<editor-fold desc="Variables for table $prefix.'dnt_members'">
		$fields_array_members	= array(
			'code',
			'code_type',
			'code_no',
			'code_year',
			'prefix',
			'name',
			'name_card',
			'position',
			'department',
			'coordinator',
			'business',
			'counselor',
			'email',
			'journaltype_id',
			'on_behalf','boi',
			'register_on',
			'edit_address',
			'numcpd','numcpa',
			'cardid',
			'state',
			'bill_prefix',
			'bill_name',
			'bill_address',
			'bill_village',
			'bill_building',
			'bill_floor',
			'bill_room',
			'bill_alley',
			'bill_road',
			'bill_subdistrict',
			'bill_district',
			'bill_province',
			'bill_postcode',
			'bill_phone',
			'bill_fax',
			'shipment_prefix',
			'shipment_name',
			'shipment_address',
			'shipment_village',
			'shipment_building',
			'shipment_floor',
			'shipment_room',
			'shipment_alley',
			'shipment_road',
			'shipment_subdistrict',
			'shipment_district',
			'shipment_province',
			'shipment_postcode',
			'shipment_phone',
			'shipment_mobile',
			'shipment_fax',
			'invoice_number',
			'comment',
			'created',
			'created_by',
			'modified',
			'renewdate',
			'expiredate',
			'paymentdate'
		);
		$fields_members			= "(".implode(',', $fields_array_members).")";
		//</editor-fold>

		//<editor-fold desc="Variables for table $prefix.'dnt_member_subscriptions'">
		$fields_array_member_subscriptions	= array(
			'member_id',
			'journal_id',
			'package_id',
			'period',
			'unitperiod',
			'numissue',
			'price',
			'discount',
			'renewdate',
			'expiredate'
		);
		$fields_member_subscriptions		= "(".implode(',', $fields_array_member_subscriptions).")";
		$member_subscriptions				= array();
		//</editor-fold>

		//<editor-fold desc="Variables for table $prefix.'dnt_members_premium_relations'">
		$premiums								= array();
		$fields_array_members_premium_relations	= array(
			'member_id',
			'subscription_id',
			'premium_id',
			'package_id',
			'num',
			'shipment_status',
			'shipment_date'
		);
		$fields_members_premium_relations		= "(".implode(',', $fields_array_members_premium_relations).")";
		$members_premium_relations				= array();
		//</editor-fold>

		// Step to each rows of data in file
		$isFail		= false; // Flag to determine if transactions of members insertion fail or not
		try {
			echo "<pre>";
			foreach($rows as $row) {
				$cols	= $row;
				// prepare variables //
				$code_year		= (string) ( (int) substr($cols['D'], -2) + 43);
				$code			= sprintf('%s-%s%05d', $cols['B'], $code_year, $cols['A']);
				$edit_address	= (int) $cols['BM'];

				$modified		= makeTime( trim( $cols['D'] ) );
				$created		= makeTime( trim( $cols['AU'] ) );
				$renew			= makeTime( trim( $cols['AW'] ) );
				$expire			= makeTime( trim( $cols['AX'] ) );
				$payment		= makeTime( trim( $cols['AY'] ) );

				$arg['sd']		= 'subdistrict';
				$arg['d']		= 'district';
				$arg['p']		= 'provinces';

				$shipment_subdistrict	= preg_split('%แขวง|ต\.%',	trim($cols['U']),	-1, PREG_SPLIT_NO_EMPTY);
				$shipment_district		= preg_split('%เขต|อ\.%',	trim($cols['V']),	-1, PREG_SPLIT_NO_EMPTY);
				$shipment_province		= preg_split('%จ\.%',		trim($cols['W']),	-1, PREG_SPLIT_NO_EMPTY);
				$bill_subdistrict		= preg_split('%แขวง|ต\.%',	trim($cols['AE']),	-1, PREG_SPLIT_NO_EMPTY);
				$bill_district			= preg_split('%เขต|อ\.%',	trim($cols['AF']),	-1, PREG_SPLIT_NO_EMPTY);
				$bill_province			= preg_split('%จ\.%',		trim($cols['AG']),	-1, PREG_SPLIT_NO_EMPTY);

				$shipment_subdistrict	= $shipment_subdistrict[0];
				$shipment_district		= $shipment_district[0];
				$shipment_province		= $shipment_province[0];
				$bill_subdistrict		= $bill_subdistrict[0];
				$bill_district			= $bill_district[0];
				$bill_province			= $bill_province[0];

				$sd_id_s	= sprintf('SELECT id FROM '.$prefix.'dnt_%1$s %1$s WHERE %1$s.name LIKE \'%%%2$s%%\'', $arg['sd'],	$shipment_subdistrict);
				$sd_id_b	= sprintf('SELECT id FROM '.$prefix.'dnt_%1$s %1$s WHERE %1$s.name LIKE \'%%%2$s%%\'', $arg['sd'],	$bill_subdistrict);
				$d_id_s		= sprintf('SELECT id FROM '.$prefix.'dnt_%1$s %1$s WHERE %1$s.name LIKE \'%%%2$s%%\'', $arg['d'],	$shipment_district);
				$d_id_b		= sprintf('SELECT id FROM '.$prefix.'dnt_%1$s %1$s WHERE %1$s.name LIKE \'%%%2$s%%\'', $arg['d'],	$bill_district);
				$p_id_s		= sprintf('SELECT id FROM '.$prefix.'dnt_%1$s %1$s WHERE %1$s.name LIKE \'%%%2$s%%\'', $arg['p'],	$shipment_province);
				$p_id_b		= sprintf('SELECT id FROM '.$prefix.'dnt_%1$s %1$s WHERE %1$s.name LIKE \'%%%2$s%%\'', $arg['p'],	$bill_province);

				$shipment_subdistrict	= $mysqli->query($sd_id_s)->fetch_object()->id;
				$bill_subdistrict		= $mysqli->query($sd_id_b)->fetch_object()->id;
				$shipment_district		= $mysqli->query($d_id_s)->fetch_object()->id;
				$bill_district			= $mysqli->query($d_id_b)->fetch_object()->id;
				$shipment_province		= $mysqli->query($p_id_s)->fetch_object()->id;
				$bill_province			= $mysqli->query($p_id_b)->fetch_object()->id;
				// ================= //

				// initialize variable for data
				$values			= array(
						'code'					=>	$code,						#
						'code_type'				=>	$cols['B'],					#
						'code_no'				=>	$cols['A'],					#int(11)
						'code_year'				=>	$code_year,					#
						'prefix'				=>	$cols['E'],					#
						'name'					=>	$cols['F'],					#
						'name_card'				=>	$cols['H'],					#
						'position'				=>	$cols['I'],					#
						'department'			=>	'',							#
						'coordinator'			=>	'',							#
						'business'				=>	$cols['AM'],				#
						'counselor'				=>	$cols['AL'],				#
						'email'					=>	$cols['BN'],				#
						'journaltype_id'		=>	(int) $cols['B'],			#int(11)
						'on_behalf'				=>	'',							#
						'boi'					=>	$cols['BO'],				#
						'register_on'			=>	'',							#
						'edit_address'			=>	$edit_address,				#tinyint(1)
						'numcpd'				=>	'',							#
						'numcpa'				=>	'',							#
						'cardid'				=>	'',							#
						'state'					=>	'',							#tinyint(3)
						'bill_prefix'			=>	$cols['L'],					#
						'bill_name'				=>	$cols['M'],					#
						'bill_address'			=>	$cols['X'],					#
						'bill_village'			=>	$cols['Y'],					#
						'bill_building'			=>	$cols['Z'],					#
						'bill_floor'			=>	$cols['AA'],				#
						'bill_room'				=>	$cols['AB'],				#
						'bill_alley'			=>	$cols['AC'],				#
						'bill_road'				=>	$cols['AD'],				#
						'bill_subdistrict'		=>	$bill_subdistrict,			#int(11)
						'bill_district'			=>	$bill_district,				#int(11)
						'bill_province'			=>	$bill_province,				#int(11)
						'bill_postcode'			=>	$cols['AB'],				#
						'bill_phone'			=>	$cols['AI'],				#
						'bill_fax'				=>	$cols['AK'],				#
						'shipment_prefix'		=>	$cols['J'],					#
						'shipment_name'			=>	$cols['K'],					#
						'shipment_address'		=>	$cols['N'],					#
						'shipment_village'		=>	$cols['O'],					#
						'shipment_building'		=>	$cols['P'],					#
						'shipment_floor'		=>	$cols['Q'],					#
						'shipment_room'			=>	$cols['R'],					#
						'shipment_alley'		=>	$cols['S'],					#
						'shipment_road'			=>	$cols['T'],					#
						'shipment_subdistrict'	=>	$shipment_subdistrict,		#int(11)
						'shipment_district'		=>	$shipment_district,			#int(11)
						'shipment_province'		=>	$shipment_province,			#int(11)
						'shipment_postcode'		=>	$cols['AQ'],				#
						'shipment_phone'		=>	$cols['AH'],				#
						'shipment_mobile'		=>	$cols['BQ'],				#
						'shipment_fax'			=>	$cols['AJ'],				#
						'invoice_number'		=>	$cols['AP'],				#
						'comment'				=>	$cols['BH'],				#
						'created'				=>	$created,					#
						'created_by'			=>	'',							#
						'modified'				=>	$modified,					#
						'renewdate'				=>	$renew,						#
						'expiredate'			=>	$expire,					#
						'paymentdate'			=>	$payment					#
					);
				$values			= "('".implode("','", $values)."')";

				// query for MySQL
				$query		= 'INSERT INTO '.$prefix.'dnt_members '.$fields_members." VALUES ".$values;

				// Set not to auto commit, commit manually
				$mysqli->autocommit(false);
				// To check if insertion is successful, commit if successful or rollback if fail
				$mysqli->query($query) ? NULL : $isFail = true;
				$member_id		= $mysqli->insert_id;

				//<editor-fold desc="store additional data for insertion to $prefix.'dnt_subscription' after successfully inserted">
				//TODO store additional data for insertion to $prefix.'dnt_member_subscriptions'
				$issue_num	= $cols['AV'];
				$period		= $issue_num / 12;
				$price		= $cols['AZ'];
				$discount	= $cols['BB'];

				$params		= array();
				$params['member_id']	= $member_id;
				$params['journal_id']	= 0;
				$params['package_id']	= 0;
				$params['period']		= $period;
				$params['unitperiod']	= 'year';
				$params['numissue']		= $issue_num;
				$params['price']		= $price;
				$params['discount']		= $discount;
				$params['renewdate']	= $renew;
				$params['expiredate']	= $expire;
				//</editor-fold>
				$member_subscriptions[]	= "('".implode("','", $params)."')";

				//TODO store additional data for insertion to $prefix.'dnt_members_premium_relations'
				$premiums[]		= $cols['BG'];
				#-----END OF ROW-----#
			}

			// check transaction when all rows stepped in to determine next process.
			if($isFail) {
				$mysqli->rollback(); echo "<h3>Insertion's failed (members). Please checks your data and try again.</h3>";
				echo "<em>( Error : ".$mysqli->errno." - ".$mysqli->error.")</em><hr/>";
			}
			else {
				// Tell for temporary insertion process complete, waiting all steps complete so main process can commit.
				echo "<h3>Insertion to table 'members' is completed.<br/>Proceeding to next step >><br/>Insertion to table 'member_subscriptions'...</h3><hr/>";

				// Begin : table 'member_subscription' insertion process
				$query_begin = 'INSERT INTO '.$prefix.'dnt_member_subscriptions '.$fields_member_subscriptions.' VALUES';
				$isFail = false; // Flag to determine if transaction of member_subscriptions insertion is fail or not
				foreach($member_subscriptions as $i => $info) {
					// Set not to auto commit, commit manually
					$mysqli->autocommit(false);
					$query	= $query_begin.$info;
					$mysqli->query($query) ? null : $isFail = true;
					$subscription_id	= $mysqli->insert_id;

					$info = preg_replace('/^\(/', '', $info);
					$info = preg_replace('/\)$/', '', $info);
					$vars	= explode(',', $info); // get data from 'member_subscription'

					// explode premium name of each rows
					$premium_names				= array();
					$premium_items				= explode(',', $premiums[$i]);
					foreach($premium_items as $item) {
						$premium					= explode('-', $item);
						$premiumName				= trim($premium[0]);

						if(preg_match('/^ปม./', $premiumName)) $premiumName = str_replace('ปม.', 'ประมวลรัษฎากร ', $premiumName);

						$premiumNum					= explode(' ', trim($premium[1]) );
						$premiumNum					= $premiumNum[0];
						$query_str					= "SELECT id FROM ".$prefix."dnt_premium p WHERE p.name LIKE '%".$premiumName."%'";
						$result						= $mysqli->query($query_str)->fetch_object();

						if(!empty($result))
							$premiumId				= $result->id;
						else
							$premiumId				= 0;

						$params						= array();
						$params['member_id']		= str_replace("'", '', $vars[0]);
						$params['subscription_id']	= $subscription_id;
						$params['premium_id']		= $premiumId;
						$params['package_id']		= 0;
						$params['num']				= $premiumNum;
						$params['shipment_status']	= 1;
						$params['shipment_date']	= 0;

						$members_premium_relations[]	= "('".implode("','", $params)."')";
					}
				}

				if($isFail) {
					$mysqli->rollback(); echo "<h3>Insertion's failed (member_subscriptions). Please checks your data and try again.</h3>";
					echo "<em>( Error : ".$mysqli->errno." - ".$mysqli->error.")</em><hr/>";
				}
				else {
					echo "<h3>Insertion to table 'member_subscriptions' is completed.<br/>Proceeding to next step >><br/>Insertion to table 'members_premium_relations'...</h3><hr/>";

					// Begin : table 'members_premium_relations' insertion process
					$query_context	= "INSERT INTO ".$prefix."dnt_members_premium_relations".$fields_members_premium_relations." VALUES";
					$isFail = false;
					foreach($members_premium_relations as $data) {
						preg_replace('/^\(/', '', $data);
						preg_replace('/\)$/', '', $data);

						// Set not to auto commit, commit manually
						$mysqli->autocommit(false);
						$query = $query_context.$data;
						$mysqli->query($query) ? null : $isFail = true;
					}

					// table 'members' checking transaction whether to commits query for real insertion.
					if($isFail) {
						$mysqli->rollback(); echo "<h3>Insertion's failed (members_premium_relations). Please checks your data and try again.</h3>";
						echo "<em>( Error : ".$mysqli->errno." - ".$mysqli->error.")</em><hr/>";
					}
					else {
						//TODO Test commit() and rollback()
						$mysqli->commit();
						echo "<h2>Transaction is committed!</h2><hr/><h2>End of Program...</h2><hr/>";
					}
				}
			}
			echo "</pre>";
			//============ END OF PROGRAM ============//
		}
		catch(Exception $e) {
			$mysqli->rollback(); echo "<h3>Insertion's failed. Please checks your data and try again.</h3>";
			echo "<em>( Error : ".$mysqli->errno." - ".$mysqli->error.")</em><hr/>";
		}
	}
	$mysqli->close();
	/* --------------------------------------------------------------------- */

	/* ---------------------------- Footer ---------------------------- */
	echo "<p>Peak memory usage: ".(memory_get_peak_usage(true) / 1024 / 1024)." MB</p>";
	$finish		= microtime(true);
	echo "<p>Finished : ".number_format($finish - $start, 2).' s</p>';
	/* ---------------------------------------------------------------- */