<?php
if (!defined('MainFile')) { header($_SERVER['SERVER_PROTOCOL'].' 404 Not Found'); die (); }

require_once PhpIncludePath."ChangeDB.inc";
require_once PhpIncludePath."MessageStrings.inc";
require_once PhpIncludePath."GlobalFunctions.php";

define ('SendStageNext_CheckCanSend', 1);
define ('SendStageNext_ProcessRTF', 2);
define ('SendStageNext_ConvertToPDF', 3);
define ('SendStageNext_ConfirmResendEmail', 4);
define ('SendStageNext_SendEmail', 5);
define ('SendStageNext_SendStage', 6);
define ('SendStageNext_Done', 7);
define ('SendStageNext_Fail', 8);

class SendStage {
	private static $SQ= NULL;
	private static $aStage= FALSE;
	private static $StageId= NULL;
	private static $SentAsInfo= NULL;
	private static $StagePath= NULL;
	private static $aStageFiles= NULL;
	private static $RecipientDetails= NULL;
	private static $ExtRecipientInfo= NULL;
	
	function __destruct() {
		if (self::$SQ)
			self::$SQ->SQL_Disconnect();
		logger::Trace('Destroying SendStage');
	}	
   
	public static function ProcessCurrentStep ($StageId, $ExtraInfo= 0) {
		if (!is_nonempty_bigint($StageId) || self::$StageId && self::$StageId !== $StageId) {
			logger::Message("ProcessCurrentStep: Incorrect StageId={$StageId}".(self::$StageId ? " while current is ".self::$StageId : ''), FakegetLogFile);
			return FALSE;
		}
		self::$StageId= $StageId;
		
		if (!isset($_SESSION['ProcessedStageStatus'])) {
			$_SESSION['ProcessedStageStatus']= SendStageNext_CheckCanSend;
			$_SESSION['ProcessngStageId']= $StageId;
			$_SESSION['ProcessngStageEmail']= '';
			$_SESSION['ProcessngStageEmailType']= 0;
		} else if ($_SESSION['ProcessedStageStatus'] == SendStageNext_CheckCanSend
								||!isset($_SESSION['ProcessngStageId'])
								|| $_SESSION['ProcessngStageId'] != $StageId) {
			logger::Message("ProcessCurrentStep: Unexpected params StageId={$StageId} (expected {$_SESSION['ProcessngStageId']}), Status= {$_SESSION['ProcessedStageStatus']}", FakegetLogFile);
			return FALSE;
		}
		$init_state= $_SESSION['ProcessedStageStatus'];

		profiler::addTag("ProcessCurrentStep Start stage {$StageId} processing from status {$init_state}, connecting DB");
		if (!self::$SQ) {
			self::$SQ= new ChangeDB();
			self::$SQ->SQL_Connect();
		}
		profiler::addTag("ProcessCurrentStep DB connected");
		if (!self::$aStage)
			self::$aStage= self::$SQ->CheckCaseHolderForStage($StageId, $_SESSION['ActualLoggedInEntityID']);

		$res= FALSE;
		while (self::$aStage !== FALSE) {
			profiler::addTag("ProcessCurrentStep ProcessedStageStatus={$_SESSION['ProcessedStageStatus']}");
			if (self::$aStage[CheckCaseHolderForStage_StageCatageory] != STC_SEND)
				return 'INCORRECT_TYPE';
			if (!intval(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_StageSendAs]))
				return 'INCORRECT_SENDAS';

			switch ($_SESSION['ProcessedStageStatus']) {
				case SendStageNext_CheckCanSend:
					$res= self::CheckCanSend();
					break;
				case SendStageNext_ProcessRTF:
					$res= self::ProcessRTF();
					break;
				case SendStageNext_ConvertToPDF:
					$res= self::ConvertToPDF();
					break;
				case SendStageNext_ConfirmResendEmail:
					$res= self::ProcessResendChoice($ExtraInfo);
					break;
				case SendStageNext_SendEmail:
					$res= self::SendStageEmail($_SESSION['ProcessngStageEmailType']);
					break;
				case SendStageNext_SendStage:
					$res= self::UpdateDB();
					break;
			}
			profiler::addTag("ProcessCurrentStep done {$init_state}=>({$_SESSION['ProcessedStageStatus']})=".($res !== FALSE ? $res : 'FALSE'));
			logger::Trace("ProcessCurrentStep done {$init_state}=>({$_SESSION['ProcessedStageStatus']})=".($res !== FALSE ? $res : 'FALSE'));
			if ($res != 'NOEMAIL')
				break;
		}
		return $res;
	}

	public static function ProcessStageFiles ($StageId, $aFiles= NULL, $SQ= NULL) {
		if (!is_nonempty_bigint($StageId)) {
			logger::Message("ProcessStageFiles: Incorrect StageId={$StageId}", FakegetLogFile);
			return FALSE;
		}
		self::$StageId= $StageId;

		if ($aFiles)
			foreach ($aFiles as $FileId)
				if (!is_bigint($FileId))
					return FALSE;

		if ($SQ) self::$SQ= $SQ;
		if (!self::$SQ) {
			self::$SQ= new ChangeDB();
			self::$SQ->SQL_Connect();
		}
		if (!self::$aStage && !(self::$aStage= self::$SQ->CheckCaseHolderForStage($StageId, $_SESSION['ActualLoggedInEntityID'])))
			return FALSE;

		if (!self::CreateFilesArray($aFiles, TRUE))
			return FALSE;

		$DestPath= MakeDirPath(array(SocketServerPath, $_SESSION['PEID'], $_SESSION['ActualLoggedInEntityID']));
		$ProcessStatus= self::ProcessStage($DestPath) && self::UpdateStageFiles();

		self::AuditStageFileProcessing(AUDIT_PROCESSFILE);
		self::ClearTempFiles($DestPath);
		if ($ProcessStatus)
			return TRUE;
		else {
			$retArray= array ();
			foreach (self::$aStageFiles[4] as $idx => $FileStatus)
				if ($FileStatus === WAPS_ProcessFail || $FileStatus === WAPS_ProcessedWithErrors)
					$retArray[]= array (self::$aStageFiles[3][$idx][0], self::$aStageFiles[0][$idx], self::$aStageFiles[4][$idx]);
			return $retArray;
		}
	}

	public static function ConvertAndSendStageEmail (&$aStage, $SQ= NULL) {
		if ($SQ) self::$SQ= $SQ;
		if (!self::$SQ) {
			self::$SQ= new ChangeDB();
			self::$SQ->SQL_Connect();
		}

		self::$aStage= $aStage;
		if (!self::$aStage)
			throw new Exception ("Incorrect stage array");
		self::$StageId= $aStage[CheckCaseHolderForStage_StageBase + Stage_CaseStageID];

		if (DBDateEmpty($aStage[CheckCaseHolderForStage_StageBase + Stage_SentOn]))
			throw new Exception ("Fail to send stage email of unprocessed stage StageId=".self::$StageId);
		if ($aStage[CheckCaseHolderForStage_StageBase + Stage_StageSendAs] == SA_QUICKEMAIL)
			throw new Exception ("Quick email have may not be send for StageId=".self::$StageId);
		if (!self::$SentAsInfo && !(self::$SentAsInfo= self::$SQ->QueryStageSendAs(self::$StageId)))
			throw new Exception ("Fail to get SentAsInfo for StageId=".self::$StageId);
		if (self::$SentAsInfo[SA_IDX_EMAIL] != 1)
			throw new Exception ("Stage StageId=<{self::$StageId}> does not require email send");
		if (!($RecipientDetails= self::$SQ->QueryRecipientNew(self::$StageId)))
			throw new Exception ("Fail to get RecipientDetails for StageId=".self::$StageId);
		if (FALSE === ($RecipientEmail= self::$SQ->CheckSendStageEmail($aStage, self::$SentAsInfo, $RecipientDetails, $EmailType, $StageEmail)))
			throw new Exception ("Quick email have may not be send for StageId=".self::$StageId);
		if (!$RecipientEmail || !filter_var($RecipientEmail, FILTER_VALIDATE_EMAIL))
			throw new Exception ("Incorrect recipient email <{$RecipientEmail}> for StageId=".self::$StageId);
		if (!$EmailType)
			throw new Exception ("Unable to select EmailType for StageId=".self::$StageId);
			
		if (!self::CreateFilesArray())
			throw new Exception ("Fail to get attachments list for StageId=".self::$StageId);

		if (FALSE !== ($idx= self::$SQ->GetReceiptPresent(self::$aStageFiles)))
			foreach (self::$aStageFiles as $i => $aList)
				unset(self::$aStageFiles[$i][$idx]);

		if (!self::ConvertFiles()) {
			self::ClearPDFFiles();
			throw new Exception ("Fail to convert files to PDF for StageId=".self::$StageId);
		}

		logger::Message("ConvertAndSendStageEmail for StageId=".self::$StageId." {$StageEmail}", MailLogFile);
		$res= self::SendStageEmail($EmailType, FALSE);
		self::ClearPDFFiles();

		if ($res === FALSE || $res === 'INCORRECT_EMAILSEND')
			throw new Exception ("Fail to send email <{$res}> for StageId=".self::$StageId);
	}

	private static function CheckCanSend() {
		$RecipientId= self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientId] && is_bigint(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientId]) ? self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientId] : 0;
		$RecipientCaseId= self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientCaseId] && is_bigint(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientCaseId]) ? self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientCaseId] : 0;
		
		if (!self::$SentAsInfo && !(self::$SentAsInfo= self::$SQ->QueryStageSendAs(self::$StageId)))
			return FALSE;
		if (!self::$RecipientDetails && !(self::$RecipientDetails= self::$SQ->QueryRecipientNew(self::$StageId)))
			return FALSE;

		$EmailType= 0;
		$StageEmail= 'NOEMAIL';
		$ret= self::$SQ->CheckCanSend(self::$aStage, self::$SentAsInfo, self::$RecipientDetails, $EmailType, $StageEmail);
		$_SESSION['ProcessngStageEmailType']= $EmailType;
		$_SESSION['ProcessngStageEmail']= $StageEmail;
		if ($ret !== TRUE)
			return $ret;

		if (!self::CreateFilesArray())
			return FALSE;
		$bRTFpresent= FALSE;
		$aDocList= array ();
		$bPDFrequired= FALSE;
		foreach (self::$aStageFiles[1] as $idx => $Ext) {
			$ExtL= strtolower($Ext);
			if ($ExtL == 'doc' || $ExtL == 'docx')
				$aDocList[]= array (self::$aStageFiles[3][$idx][0], self::$aStageFiles[0][$idx], WAPS_NotProcessed);
			if ($ExtL == 'rtf')
				$bRTFpresent= TRUE;
			if (!is_file(self::$aStageFiles[2][$idx])) {
				logger::Error("CheckCanSend detect missing stage file ".self::$aStageFiles[2][$idx]);
				return 'INCORRECT_FILE';
			}
			if (self::NeedToConvertToPdf($idx))
				$bPDFrequired= TRUE;
		}

		$bRTFpresent= $bRTFpresent && DBDateEmpty(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_SentOn]);
		if ($bRTFpresent)
			$_SESSION['ProcessedStageStatus']= SendStageNext_ProcessRTF;
		elseif ($bPDFrequired)
			$_SESSION['ProcessedStageStatus']= SendStageNext_ConvertToPDF;
		elseif (strpos($_SESSION['ProcessngStageEmail'], 'EMAILAGAIN') === 0)
			$_SESSION['ProcessedStageStatus']= SendStageNext_ConfirmResendEmail;
		else
			$_SESSION['ProcessedStageStatus']= SendStageNext_SendEmail;

		if (count($aDocList) && DBDateEmpty(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_SentOn]))
			return $aDocList;
		else if ($bRTFpresent)
			return 'PROCESSING';
		else if ($bPDFrequired)
			return 'CONVERTING';
		else
			return $_SESSION['ProcessngStageEmail'];
	}
	
	private static function NeedToConvertToPdf ($idx) {
		$ConvertableExt= array ('doc', 'docx', 'rtf');

		if (!self::CreateFilesArray())
			return FALSE;

		if (!isset(self::$aStageFiles[1][$idx]) || !isset(self::$aStageFiles[3][$idx]))
			return FALSE;
		$ExtL= strtolower(self::$aStageFiles[1][$idx]);
		return in_array($ExtL, $ConvertableExt) && 
						(self::$aStageFiles[3][$idx][0] == 0 || self::$aStageFiles[3][$idx][6] == 'N');
	}
	
	private static function GetStagePath () {
		if (self::$StagePath) return;
		self::$StagePath= MakeDirPath(array(OWCClientFilesPath,  self::$aStage[CheckCaseHolderForStage_CaseHolderFirmID], self::$aStage[CheckCaseHolderForStage_ModuleID], self::$aStage[CheckCaseHolderForStage_SubModule], self::$aStage[CheckCaseHolderForStage_CaseHolderID], self::$aStage[CheckCaseHolderForStage_StageBase + Stage_CaseID], 'S'));
	}
	
	private static function CreateFilesArray ($aFiles= NULL, $bOnlyProcessable= FALSE) {
		if (self::$aStageFiles) return TRUE;
		
		self::GetStagePath();
		self::$aStageFiles= array (array (), array (), array (), array ());
		$AttachCount= self::$SQ->CreateStageFilesArray(self::$aStage, self::$aStageFiles, TRUE);
		// Filter files array
		if ($AttachCount !== FALSE && ($aFiles || $bOnlyProcessable)) {
			$aFilteredStageFiles= array (array (), array (), array (), array ());
			foreach (self::$aStageFiles[3] as $idx => $AttachId) {
				if (($aFiles && in_array($AttachId[0], $aFiles)	|| !$aFiles)
							&& ($bOnlyProcessable && strtolower(self::$aStageFiles[1][$idx]) == 'rtf' || !$bOnlyProcessable))
					for ($i= 0; $i < 4; $i++)
						$aFilteredStageFiles[$i][]= self::$aStageFiles[$i][$idx];
			}
			self::$aStageFiles= $aFilteredStageFiles;
		}
		return $AttachCount !== FALSE;
	}
	
	private static function ProcessRTF () {
		if (!self::CreateFilesArray())
			return FALSE;

		$DestPath= MakeDirPath(array(SocketServerPath, $_SESSION['PEID'], $_SESSION['ActualLoggedInEntityID']));
		$ProcessStatus= self::ProcessStage($DestPath) && self::UpdateStageFiles();
		self::AuditStageFileProcessing(AUDIT_PROCESSFILE);
		self::ClearTempFiles($DestPath);

		if ($ProcessStatus) {
			$bPDFrequired= FALSE;
			foreach (self::$aStageFiles[1] as $idx => $Ext)
				if (self::NeedToConvertToPdf($idx)) {
					$bPDFrequired= TRUE;
					break;
				}
			if ($bPDFrequired) {
				$_SESSION['ProcessedStageStatus']= SendStageNext_ConvertToPDF;
				return 'CONVERTING';
			} else {
				if (strpos($_SESSION['ProcessngStageEmail'], 'EMAILAGAIN') === 0)
					$_SESSION['ProcessedStageStatus']= SendStageNext_ConfirmResendEmail;
				else
					$_SESSION['ProcessedStageStatus']= SendStageNext_SendEmail;
				return $_SESSION['ProcessngStageEmail'];
			}
		} else {
			$_SESSION['ProcessedStageStatus']= SendStageNext_Fail;
			self::SetStageAsNotSent();
			$retArray= array ();
			foreach (self::$aStageFiles[4] as $idx => $FileStatus)
				if ($FileStatus === WAPS_ProcessFail || $FileStatus === WAPS_ProcessedWithErrors)
					$retArray[]= array (self::$aStageFiles[3][$idx][0], self::$aStageFiles[0][$idx], self::$aStageFiles[4][$idx]);
			return $retArray;
		}
	}
	
	private static function RemoveAcknowledgementOfReceipt() {
		if (!self::CreateFilesArray())
			return FALSE;

		$ReceiptPresent= self::$SQ->GetReceiptPresent(self::$aStageFiles);
		if ($ReceiptPresent === FALSE)
			return TRUE;

		$SubModuleId= self::$aStage[CheckCaseHolderForStage_SubModule];
		$CaseId= self::$aStage[CheckCaseHolderForStage_StageBase + Stage_CaseID];

		self::$SQ->SQL_BeginTransaction("Delete acknowledgment of Receipt in the stage ".self::$StageId." - 1 Query");
		$aFilesList= self::$SQ->deleteAttachments(self::$aStage, array (self::$aStageFiles[3][$ReceiptPresent][0]), self::$aStageFiles);
		if ($aFilesList) {
			self::$SQ->SQL_CommitTransaction();
			self::$SQ->AuditFileChanges($CaseId, $CaseId, $aFilesList);
			$ret= TRUE;
		} else {
			self::$SQ->InsertAuditWarning($CaseId, $CaseId, array (AUDIT_ATTACHFILE), NULL, self::$StageId, AcknowledgementOfReceiptName);
			$ret= FALSE;
		}
		for ($j= 0, $l= count(self::$aStageFiles); $j < $l; $j++)
			unset(self::$aStageFiles[$j][$ReceiptPresent]);

		return $ret;
	}

	private static function SetStageAsNotSent () {
		$SubModuleId= self::$aStage[CheckCaseHolderForStage_SubModule];
		$CaseId= self::$aStage[CheckCaseHolderForStage_StageBase + Stage_CaseID];

		self::$SQ->SetStageAsNotSent($CaseId, $SubModuleId, self::$StageId);
	}

	private static function ConvertToPDF () {
		if (!self::CreateFilesArray())
			return FALSE;

		logger::addLogLevel(LogAll);

		self::RemoveAcknowledgementOfReceipt();
		profiler::addTag("ConvertToPDF started");
		$ConvertStatus= self::ConvertFiles() && self::UpdateStageFiles();
		profiler::addTag("ConvertToPDF convert completed with status ".($ConvertStatus ? 'SUCCESS' : 'FAIL'));
		self::AuditStageFileProcessing(AUDIT_CONVERTFILEPDF);
		profiler::addTag("ConvertToPDF audited");
		self::ClearPDFFiles();
		profiler::addTag("ConvertToPDF cleared");

		if ($ConvertStatus) {
			if (strpos($_SESSION['ProcessngStageEmail'], 'EMAILAGAIN') === 0)
				$_SESSION['ProcessedStageStatus']= SendStageNext_ConfirmResendEmail;
			else
				$_SESSION['ProcessedStageStatus']= SendStageNext_SendEmail;
			return $_SESSION['ProcessngStageEmail'];
		} else {
			$_SESSION['ProcessedStageStatus']= SendStageNext_Fail;
			self::SetStageAsNotSent();
			return FALSE;
		}
	}

	private static function ConvertFileLocally (&$socket, &$Word, $idx, &$TmpPdfFile) {
		profiler::addTag("ConvertFilesLocally converting ".self::$aStageFiles[2][$idx]." , starting OpenSocket");
		$socket->OpenSocket();
		if (!$socket->SocketExist())
			throw new Exception ('fail to create socket');
		$TmpPdfFile= TmpDir.'pdftmp'.round(microtime(true)).'.'.self::$aStageFiles[1][$idx];
		profiler::addTag("ConvertFilesLocally OpenSocket done, creating tmp file");
		if (!copy(self::$aStageFiles[2][$idx], $TmpPdfFile))
			throw new Exception ('Fail to create temp file');
		profiler::addTag("ConvertFilesLocally opening tmp file {$TmpPdfFile}");
		$Word->Open($TmpPdfFile);
		profiler::addTag("ConvertFilesLocally tmp file opened, setting printer");
		$Word->SetActivePrinter(PDFprinterName);
		profiler::addTag("ConvertFilesLocally tmp file opened, checking printer");
		if (($CurPrinter= $Word->GetActivePrinter()) != PDFprinterName)
			throw new Exception ('Fail to setup printer '.PDFprinterName.' current is '.$CurPrinter.' for file '.self::$aStageFiles[2][$idx]);
		profiler::addTag("ConvertFilesLocally printer OK, starting print loop");
		for ($i= 0; $i < 5; $i++) {
			profiler::addTag("ConvertFilesLocally starting print ".self::$aStageFiles[2][$idx]." attempt #{$i}");
			$Word->PrintDocument();
			profiler::addTag("ConvertFilesLocally print complete, start ListenPdfRequest");
			if (($PDFfileName= $socket->ListenPdfRequest()) && $PDFfileName != 'ERROR')
				break;
		}
		if (!$PDFfileName || $PDFfileName == 'ERROR')
			throw new Exception ("Fail to get name of generated PDF after {$i} retries");
		profiler::addTag("ConvertFilesLocally received PDFfileName={$PDFfileName}, closing socket");
		$socket->CloseSocket();
		profiler::addTag("ConvertFilesLocally socket closed, closing tmp file in Word");
		$Word->Close();
		self::$aStageFiles[4][$idx]= WAPS_ConvertedOK;
		self::$aStageFiles[5][$idx]= PDFSpoolPath.$PDFfileName.'.pdf';
		profiler::addTag("ConvertFilesLocally tmp file released by Word, removing it");
		unlink($TmpPdfFile);
		profiler::addTag("ConvertFilesLocally tmp file removed");
	}

	private static function ConvertFiles () {
		if (!self::CreateFilesArray())
			return FALSE;

		$bPDFrequired= FALSE;
		foreach (self::$aStageFiles[1] as $idx => $Ext)
			if (self::NeedToConvertToPdf($idx)) {
				$bPDFrequired= TRUE;
				break;
			}

		$FilesCount= count(self::$aStageFiles[1]);
		self::$aStageFiles[4]= array_fill(0, $FilesCount, WAPS_NotProcessed); // Processing status 
		self::$aStageFiles[5]= array_fill(0, $FilesCount, ''); // Temp File Path

		if (!$bPDFrequired)
			return TRUE;

		if (defined('PdfConverterScript'))
			return self::ConvertPStoPDF();
		elseif (defined('RemoteWordUrl'))
			return self::ConvertFilesRemotely();
		else
			return self::ConvertFilesLocally();
	}

	private static function ConvertPStoPDF () {
		foreach (self::$aStageFiles[2] as $idx => $FilePath)
			if (self::NeedToConvertToPdf($idx)) {
				$FileName= basename($FilePath, self::$aStageFiles[1][$idx]);
				$BaseFilePath= MakeDirPath(array(SocketServerPath, $_SESSION['PEID'], $_SESSION['ActualLoggedInEntityID']));
				if (!is_file($BaseFilePath.$FileName.'ps'))
					return FALSE;
				$PDFfileName= $FileName.'_'.$_SESSION['ActualLoggedInEntityID'].'_'.date('M-d-Y_H-i-s');
				$PdfConverterScriptOut= array ();
				$PdfConverterScriptStatus= NULL;
				logger::Trace("PdfConverterScript call: ".PdfConverterScript.$BaseFilePath.$FileName.'ps '.PDFSpoolPath.$PDFfileName.'.pdf');
				exec(PdfConverterScript.$BaseFilePath.$FileName.'ps '.PDFSpoolPath.$PDFfileName.'.pdf', $PdfConverterScriptOut, $PdfConverterScriptStatus);
				unlink($BaseFilePath.$FileName.'ps');
				logger::Trace("PdfConverterScript status for {$FileName} is {$PdfConverterScriptStatus}, full output:\n".implode ("\n", $PdfConverterScriptOut));
				if (($PdfConverterScriptStatus != 0)) return FALSE;
				self::$aStageFiles[4][$idx]= WAPS_ConvertedOK;
				self::$aStageFiles[5][$idx]= PDFSpoolPath.$PDFfileName.'.pdf';
			}
		return TRUE;
	}

	private static function ConvertFilesLocally () {
		try {
			profiler::addTag("ConvertFilesLocally execution started, creating WordOLE");
			$Word = new WordOLE();
			profiler::addTag("ConvertFilesLocally WordOLE created, creating Respond");
			$socket= new Respond(PDFrespondPort);
			profiler::addTag("ConvertFilesLocally Respond created, starting file loop");
			$TmpPdfFile= NULL;

			$PDFcount= 0;
			foreach (self::$aStageFiles[2] as $idx => $FilePath)
				if (self::NeedToConvertToPdf($idx))
					self::ConvertFileLocally($socket, $Word, $idx, $TmpPdfFile);

			profiler::addTag("ConvertFilesLocally file loop ended, closing Word");
			$Word->Quit();
			profiler::addTag("Word closed");
		  return TRUE;
		} catch (Exception $e) {
			profiler::addTag('ConvertFilesLocally exception: '.$e->getMessage());
			logger::Error('ConvertFilesLocally exception: '.$e->getMessage());
			try {
				if (isset($Word)) $Word->Quit();
			} catch (Exception $e) {
				logger::Error('ConvertFilesLocally exception: '.$e->getMessage());
			}
			if (isset($socket)) $socket->CloseSocket();
			if ($TmpPdfFile && file_exists($TmpPdfFile)) unlink($TmpPdfFile);
			return FALSE;
	  }
	}

	private static function ConvertFilesRemotely () {
		try {
			profiler::addTag("ConvertFilesRemotely");

			$PDFcount= 0;
			foreach (self::$aStageFiles[2] as $idx => $FilePath)
				if (self::NeedToConvertToPdf($idx))
					self::ConvertFileRemotely($idx);

			profiler::addTag("ConvertFilesRemotely file loop ended");
		  return TRUE;
		} catch (Exception $e) {
			profiler::addTag('ConvertFilesRemotely exception: '.$e->getMessage());
			logger::Error('ConvertFilesRemotely exception: '.$e->getMessage());
			return FALSE;
	  }
	}

	private static function ConvertFileRemotely ($idx, $Timeout= CurlDefaultTimeout) {
		$cur_option= array(	CURLOPT_HEADER => false,
                 				CURLOPT_TIMEOUT => $Timeout,
                 				CURLOPT_USERAGENT => 'Papirus Stage Processor',
                 				CURLOPT_RETURNTRANSFER => true,
                 				CURLOPT_POST => true,
                 				CURLOPT_FAILONERROR => 1
		                	);

		$cur_option[CURLOPT_POSTFIELDS]= array (	'Type'				=> 'pdf',
																							'PrintTo'			=> PDFprinterName,
																							'RespondTo'		=> PDFrespondPort,
																							'DocToPrint'	=> '@'.self::$aStageFiles[2][$idx]
																						);
		$cur_option[CURLOPT_URL]= RemoteWordUrl;
	
		profiler::addTag("ConvertFilesRemotely convertion start");
		$ch= curl_init();
		curl_setopt_array($ch, $cur_option);
		$sResult= curl_exec($ch);
	
		$aInfo= curl_getinfo ($ch);
		logger::Trace("ConvertFilesRemotely curl results: <{$sResult}>, full info:\n".print_r($aInfo, TRUE));
		$CURL_error= curl_error($ch);
		if ($CURL_error)
			throw new Exception ('CURL Error: '.$CURL_error);
	
		curl_close($ch);

		$aResults= explode(',', $sResult);
		if (!$sResult || count($aResults) != 2 || $aResults[0] != 'SUCCESS')
			throw new Exception ("Fail to get name of generated PDF");
		profiler::addTag("ConvertFilesRemotely received PDFfileName={$aResults[1]}");

		self::$aStageFiles[4][$idx]= WAPS_ConvertedOK;
		self::$aStageFiles[5][$idx]= PDFSpoolPath.$aResults[1].'.pdf';
	}

	private static function ClearTempFiles ($TmpPath) {
		$FilesList= glob($TmpPath.TempFilePrefix.'*');
		foreach ($FilesList as $file)
			unlink($file);
	}
	
	private static function ClearPDFFiles () {
		if (!is_array(self::$aStageFiles[5]))
			return FALSE;
		foreach (self::$aStageFiles[5] as $idx => $PDFfileName)
			if ($PDFfileName && is_file($PDFfileName))
				unlink($PDFfileName); 
	}

	private static function ProcessStage ($TmpPath) {
		logger::addLogLevel(LogAll);
		logger::Trace("ProcessStage ".self::$StageId." to ".$TmpPath);

		if (!self::CreateFilesArray())
			return FALSE;

		if (!($FilesCount= count(self::$aStageFiles[3]))) {
			self::$aStageFiles[4]= array ();
			self::$aStageFiles[5]= array ();
			return TRUE;
		}
		self::$aStageFiles[4]= array_fill(0, $FilesCount, WAPS_NotProcessed); // Processing status 
		self::$aStageFiles[5]= array_fill(0, $FilesCount, ''); // Temp File Path

		$WA= new RtfSimpleAutomation (self::$StageId, self::$SQ, TRUE);
		$aOutputs= array ();
		$FailCount= 0;
		for ($idx= 0; $idx < $FilesCount; $idx++) {
			$WA->selectFile(self::$aStageFiles[3][$idx][0]);
			if (!$WA->canProcessFile()) {
				logger::Trace("Unable to process file ".self::$StageId."_".self::$aStageFiles[3][$idx][0].".".self::$aStageFiles[1][$idx]);
				continue;
			}
			self::$aStageFiles[5][$idx]= tempnam($TmpPath, TempFilePrefix);
			$aOutputs[0]= new FileOutputStream (self::$aStageFiles[5][$idx]);
			if (!$WA->ProcessFile($aOutputs)) {
				logger::Error("Fail to process file ".self::$StageId."_".self::$aStageFiles[3][$idx][0].".".self::$aStageFiles[1][$idx]);
				self::$aStageFiles[4][$idx]= WAPS_ProcessFail;
				$FailCount++;
			} else {
				if ($WA->CodeErrorsCount) {
					logger::Trace("File ".self::$StageId."_".self::$aStageFiles[3][$idx][0].".".self::$aStageFiles[1][$idx]." processed with Code Error(s)");
					self::$aStageFiles[4][$idx]= WAPS_ProcessedWithErrors;
					$FailCount++;
				} else {
					logger::Trace("File ".self::$StageId."_".self::$aStageFiles[3][$idx][0].".".self::$aStageFiles[1][$idx]." processed OK");
					self::$aStageFiles[4][$idx]= WAPS_ProcessedOK;
				}
			}
			$aOutputs[0]->close();

			if (self::$aStageFiles[4][$idx] == WAPS_ProcessedOK) {
				$ColorFreeFile= tempnam($TmpPath, TempFilePrefix);
				$aOutputs[0]= new FileOutputStream ($ColorFreeFile);
				$ColorRemover= new RtfRemoveColors(self::$aStageFiles[5][$idx]);
				$RemoveResult= $ColorRemover->ProcessFile($aOutputs);
				$aOutputs[0]->close();
				if ($RemoveResult) {
					unlink(self::$aStageFiles[5][$idx]);
					self::$aStageFiles[5][$idx]= $ColorFreeFile;
				} elseif (is_file($ColorFreeFile))
					unlink($ColorFreeFile);
			}
		}
		return $FailCount <= 0;
	}

	private static function BackupStageFile ($SrcPath) {
		if (!($NamePos= strrpos($SrcPath, '/'))) {
			logger::Error("BackupStageFile: Fail to backup file {$SrcPath}");
			return FALSE;
		}
		$BackDir= substr($SrcPath, 0, $NamePos).'/bak';
		if (!is_dir($BackDir) && !mkdir($BackDir)) {
			logger::Error("BackupStageFile: Fail to create backup dir {$BackDir} for {$SrcPath}");
			return FALSE;
		}
		return copy($SrcPath, $BackDir.'/'.substr($SrcPath, $NamePos + 1));
	}

	private static function UpdateStageFiles () {
		$StageFileCount= isset(self::$aStageFiles[4]) ? count(self::$aStageFiles[4]) : 0;
		if (!$StageFileCount) 
			return TRUE;
		// Check for fail processing
		foreach (self::$aStageFiles[4] as $idx => $FileStatus)
			if ($FileStatus === WAPS_ProcessFail)
				return FALSE;
		// Rename Original stage files
		$RenameOriginalsOk= TRUE;
		foreach (self::$aStageFiles[4] as $idx => $FileStatus) {
			if ($FileStatus === WAPS_NotProcessed)
				continue;
			self::BackupStageFile(self::$aStageFiles[2][$idx]);
			if (!rename(self::$aStageFiles[2][$idx], self::$aStageFiles[2][$idx].'tmp')) {
				$RenameOriginalsOk= FALSE;
				logger::Error("UpdateStageFiles: fail to rename ".self::$aStageFiles[2][$idx]);
				break;
			}
		}
		if (!$RenameOriginalsOk) { // Fail to rename originals, return back
			for ($i= 0; $i < $idx; $i++)
				rename(self::$aStageFiles[2][$idx].'tmp', self::$aStageFiles[2][$idx]);
			return FALSE;
		}

		// Move processed files in stage folder
		$MoveProcessedOk= TRUE;
		$aMovedFiles= array ();
		$bPDFConvertion= FALSE;
		$aNewStageFileName= array ();
		foreach (self::$aStageFiles[4] as $idx => $FileStatus) {
			if ($FileStatus === WAPS_NotProcessed)
				continue;
			logger::Trace("UpdateStageFiles: move processed ".self::$aStageFiles[5][$idx]." to ".self::$aStageFiles[2][$idx]);
			$aNewStageFileName[$idx]= self::$aStageFiles[2][$idx];
			if ($FileStatus === WAPS_ConvertedOK) {
				$bPDFConvertion= TRUE;
				if (self::$aStageFiles[1][$idx])
					$aNewStageFileName[$idx]= substr($aNewStageFileName[$idx], 0, strlen($aNewStageFileName[$idx]) - strlen(self::$aStageFiles[1][$idx]) - 1);
				$aNewStageFileName[$idx].= '.pdf';
			}
			if (is_file($aNewStageFileName[$idx])) {
				logger::Trace("UpdateStageFiles: Deleting existing stage file {$aNewStageFileName[$idx]} during convertion");
				unlink($aNewStageFileName[$idx]);
			}
			if (!copy(self::$aStageFiles[5][$idx], $aNewStageFileName[$idx])) {
				$MoveProcessedOk= FALSE;
				logger::Error("UpdateStageFiles: fail to copy processed ".self::$aStageFiles[5][$idx]." to ".self::$aStageFiles[2][$idx]);
				break;
			}
			unlink(self::$aStageFiles[5][$idx]);
			$aMovedFiles[]= self::$aStageFiles[3][$idx][0];
			self::$aStageFiles[4][$idx]= WAPS_UpdatedOK;
		}
	
		if ($MoveProcessedOk && count($aMovedFiles) && $bPDFConvertion)
			$MoveProcessedOk= self::$SQ->ChangeToPdfInDatabase(self::$StageId, $aMovedFiles);

		if (!$MoveProcessedOk) { // Fail to move processed files to stage
			foreach (self::$aStageFiles[4] as $idx => $FileStatus) {
				if ($FileStatus === WAPS_NotProcessed)
					continue;
				if (is_file($aNewStageFileName[$idx]))
					unlink($aNewStageFileName[$idx]);
				copy(self::$aStageFiles[2][$idx].'tmp', self::$aStageFiles[2][$idx]);
				unlink(self::$aStageFiles[2][$idx].'tmp');
				self::$aStageFiles[4][$idx]= WAPS_UpdateFail;
			}
			return FALSE;
		}

		foreach (self::$aStageFiles[4] as $idx => $FileStatus) {
			if ($FileStatus === WAPS_NotProcessed)
				continue;
			unlink(self::$aStageFiles[2][$idx].'tmp');
		}
		logger::Trace("UpdateStageFiles: completed OK");
		return TRUE;
	}
	
	private static function AuditStageFileProcessing ($ActivityType) {
		if (!isset(self::$aStageFiles[4])) return;
		$aFilesList= array ();

		foreach (self::$aStageFiles[4] as $idx => $FileStatus) {
			if ($FileStatus === WAPS_NotProcessed)
				continue;
			$aFilesList[]= array (self::$StageId.'_'.self::$aStageFiles[3][$idx][0], $ActivityType, $FileStatus === WAPS_UpdatedOK);
		}
		self::$SQ->AuditFileChanges(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_CaseID], self::$aStage[CheckCaseHolderForStage_SubModule], $aFilesList);
	}

	private static function ProcessResendChoice ($Choice) {
		if ($_SESSION['LoggedInFirmAdmin'] && $Choice == 1
					&& strpos($_SESSION['ProcessngStageEmail'], 'EMAILAGAIN') === 0)
			$_SESSION['ProcessngStageEmail']= 'EMAIL'.substr($_SESSION['ProcessngStageEmail'], strlen('EMAILAGAIN'));
		else
			$_SESSION['ProcessngStageEmail']= 'NOEMAIL';

		$_SESSION['ProcessedStageStatus']= SendStageNext_SendEmail;
		return $_SESSION['ProcessngStageEmail'];
	}

	private static function PrepareEmailParams ($RecipientId, $RecipientCaseId, $EmailType) {
		$ParametrsArray= array ();

		self::$SQ->SetSenderEmailArgs($ParametrsArray);

		$ParametrsArray['$ReciptientFullName']= self::$ExtRecipientInfo[2];
		$ParametrsArray['$SenderFullName']= self::$ExtRecipientInfo[19];
		$ParametrsArray['$SenderFirmName']= self::$ExtRecipientInfo[6];
		$ParametrsArray['$SenderFirmAF']= self::$ExtRecipientInfo[18];
		$ParametrsArray['$SenderEmail']= self::$ExtRecipientInfo[0];
		$ParametrsArray['$SenderCaseNumber']= self::$ExtRecipientInfo[1];
		$ParametrsArray['$Subject']= self::$ExtRecipientInfo[12];
		$ParametrsArray['$LoginEntityId']= $RecipientId;
		$ParametrsArray['$LoginCaseId']= $RecipientCaseId;
		$ParametrsArray['$EmailCaseId']= self::$aStage[CheckCaseHolderForStage_StageBase + Stage_CaseID];
		$ParametrsArray['$EmailStageId']= self::$StageId;

		if (self::$SentAsInfo[SA_IDX_ID] == SA_QUICKEMAIL) {
			$ParametrsArray['$StageName']= self::$aStage[CheckCaseHolderForStage_StageBase + Stage_CaseStageName];
			$ParametrsArray['$Content']= self::$SQ->QuickEmailBody(self::$StageId);
		} else {
			if (!self::$RecipientDetails && !(self::$RecipientDetails= self::$SQ->QueryRecipientNew(self::$StageId)))
				return FALSE;
			if (self::$RecipientDetails[0] == "Caseholder") {
				$ParametrsArray['$CaseHolderFirstName']= self::$ExtRecipientInfo[15];
				$ParametrsArray['$CaseHolderMiddleName']= self::$ExtRecipientInfo[16];
				$ParametrsArray['$CaseHolderLastName']= self::$ExtRecipientInfo[17];
				$ParametrsArray['$RepresentativeType']= self::$RecipientDetails[2];
				$ParametrsArray['$CaseNumber']= self::$RecipientDetails[3];
				$ParametrsArray['$RefferedClients']= '';
				$aRecipientClients= self::$SQ->QueryRelationships2(self::$aStage[CheckCaseHolderForStage_TransactionID], self::$RecipientDetails[10], self::$RecipientDetails[8], "PHP");
				foreach ($aRecipientClients as $aClientDetails)
					$ParametrsArray['$RefferedClients'].= ($ParametrsArray['$RefferedClients'] ? ' & ' : '').$aClientDetails[1];
			} elseif ($EmailType == EMAIL_SP_Client_SendEmail) {
				$ParametrsArray['$ClientDescription']= self::$aStage[CheckCaseHolderForStage_StageBase + Stage_DescriptionToClient];
			}
		}

		return $ParametrsArray;
	}

	private static function SendStageEmail ($EmailType, $bEmailOnConfirm= TRUE) {
		if ($bEmailOnConfirm) {
			$_SESSION['ProcessedStageStatus']= SendStageNext_SendStage;
			self::RemoveAcknowledgementOfReceipt();

			if ($_SESSION['ProcessngStageEmail'] == 'NOEMAIL')
				return $_SESSION['ProcessngStageEmail'];
			if (!$EmailType)
				return FALSE;
		}

		$RecipientId= self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientId] && is_bigint(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientId]) ? self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientId] : 0;
		$RecipientCaseId= self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientCaseId] && is_bigint(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientCaseId]) ? self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientCaseId] : 0;
		if ($RecipientId && !$RecipientCaseId && $aPossibleViewers= self::$SQ->QueryCaseClientsAndContacts(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_CaseID]))
			foreach ($aPossibleViewers as $ViewerInfo)
				if ($ViewerInfo[0] == $RecipientId) {
					$RecipientCaseId= self::$aStage[CheckCaseHolderForStage_StageBase + Stage_CaseID];
					break;
				}

		if (!self::$SentAsInfo && !(self::$SentAsInfo= self::$SQ->QueryStageSendAs(self::$StageId)))
			return FALSE;

		if (!self::$RecipientDetails && !(self::$RecipientDetails= self::$SQ->QueryRecipientNew(self::$StageId)))
			return FALSE;

		if (!self::$ExtRecipientInfo && !(self::$ExtRecipientInfo= self::$SQ->QueryExternalRecipient(self::$StageId)))
			return FALSE;

		$RecipientEmail= self::$RecipientDetails[5];

		$bSendAttachments= self::$SentAsInfo[SA_IDX_EMAIL] == 1 || self::$RecipientDetails[0] != "Caseholder";

		if ($bSendAttachments) {	
			if (!self::CreateFilesArray())
				return FALSE;
			$aAttachments= array ();		
			foreach (self::$aStageFiles[3] as $idx => $AttachId) {
				if ($AttachId[0] && is_bigint($AttachId[0]))
					$FileName= self::$aStageFiles[0][$idx];
				else if ($AttachId[0] == 0) // CoverLetter
					$FileName= $GLOBALS['CoverLetterText'].": ".self::$aStageFiles[0][$idx];
				else
					continue;
				if ($bEmailOnConfirm || !isset(self::$aStageFiles[4][$idx]) || self::$aStageFiles[4][$idx] != WAPS_ConvertedOK)
					$aAttachments[]= array (self::$aStageFiles[2][$idx], $FileName.(self::$aStageFiles[1][$idx] !== "" ? ".".self::$aStageFiles[1][$idx] : ""));
				else
					$aAttachments[]= array (self::$aStageFiles[5][$idx], $FileName.".pdf");
			}
		} else
			$aAttachments= NULL;

		$ParametrsArray= self::PrepareEmailParams($RecipientId, $RecipientCaseId, $EmailType);

		if (self::$aStage[CheckCaseHolderForStage_EmailFromEntityID] &&
				($aEmailFromEntityDetails= self::$SQ->QuetyEntityInfo(self::$aStage[CheckCaseHolderForStage_EmailFromEntityID])) &&
				$aEmailFromEntityDetails[2] &&
				filter_var($aEmailFromEntityDetails[2], FILTER_VALIDATE_EMAIL)) {
			$FromEmail= $aEmailFromEntityDetails[2];
			$ParametrsArray['$EmailFromName']= $aEmailFromEntityDetails[1];
		} else {
			$FromEmail= Email_FromEmail;
			$ParametrsArray['$EmailFromName']= $_SESSION['PEN'];
		}

		if ($EmailDetails= self::$SQ->CreateEmail($EmailType, $RecipientEmail, $ParametrsArray, $aAttachments, $FromEmail, self::$SQ->GetFirmBCCList(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_CaseID], self::$aStage))) {
			return self::CreateQuickEmailCover($EmailDetails);
		}
		logger::Error("SendStageEmail returns INCORRECT_EMAILSEND, EmailDetails=".print_r($EmailDetails, TRUE));
		return 'INCORRECT_EMAILSEND';
	}

	private static function CreateQuickEmailCover ($EmailDetails) {
		if (!self::$SentAsInfo) return FALSE;
		if (self::$SentAsInfo[SA_IDX_ID] != SA_QUICKEMAIL) return TRUE;

		$EmailCreatedOK= FALSE;
		do {
			$CoverExt= 'eml';
			$CoverName= 'Email Message';
			if (!self::CreateFilesArray())
				break;
			if (!count(self::$aStageFiles[3]) || self::$aStageFiles[3][0])
				foreach (self::$aStageFiles as $arr_idx => $arr)
					array_unshift(self::$aStageFiles[$arr_idx], NULL);
			self::$aStageFiles[0][0]= $CoverName;
			self::$aStageFiles[1][0]= $CoverExt;
			self::$aStageFiles[2][0]= self::$StagePath.self::$StageId.($CoverExt !== "" ? ".".$CoverExt : "");
			self::$aStageFiles[3][0]= 0;
			if (isset(self::$aStageFiles[4]))
				self::$aStageFiles[4][0]= WAPS_NotProcessed;
			if (isset(self::$aStageFiles[5]))
				self::$aStageFiles[5][0]= '';
		
			if (!file_put_contents(self::$aStageFiles[2][0], $EmailDetails['Headers'].$EmailDetails['Body']))
				break;

			if (self::$aStage[CheckCaseHolderForStage_StageBase + Stage_DescriptionToClient] == SystemQueries::GetDefaultClientStageDescription()
					|| self::$aStage[CheckCaseHolderForStage_StageBase + Stage_DescriptionToClient] == SystemQueries::GetClientStageDescription(
																					self::$aStage[CheckCaseHolderForStage_StageBase + Stage_StageType], 
																					self::$aStage, 
																					is_nonempty_bigint(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_ParentId]) ? self::$SQ->CheckCaseHolderForStage(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_ParentId], $_SESSION['ActualLoggedInEntityID']) : NULL,
																					in_array(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_ProcedureType], array (AT_RECCALL, AT_MAKECALL, AT_RECCALL, AT_MAKECALL))
																																																																			)
					)
				$Description= self::$SQ->CreateEmailClientDescription($EmailDetails['Headers'], $EmailDetails['Body']);
			else
				$Description= NULL;

			if (!self::$SQ->InsertQuickEmailCover(self::$StageId, $CoverName, $CoverExt, $Description))
				break;
			self::$aStage[CheckCaseHolderForStage_StageBase + Stage_CoverLetterName]= self::$aStageFiles[0][0];
			self::$aStage[CheckCaseHolderForStage_StageBase + Stage_CoverLetterType]= self::$aStageFiles[1][0];
			$EmailCreatedOK= TRUE;
		} while (FALSE);
		self::$SQ->AuditFileChanges(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_CaseID], self::$aStage[CheckCaseHolderForStage_SubModule], array (array (self::$StageId, AUDIT_GENERATEFILE, $EmailCreatedOK)));
		return $EmailCreatedOK;
	}
	
	private static function UpdateDB () {
		if (!self::$RecipientDetails && !(self::$RecipientDetails= self::$SQ->QueryRecipientNew(self::$StageId)))
			return FALSE;

		self::$SQ->SQL_BeginTransaction("SendStage ".self::$StageId);
		if (!self::UpdateSendStage())
			return FALSE;

		self::$SQ->SQL_CommitTransaction();
		$_SESSION['ProcessedStageStatus']= SendStageNext_Done;
		return 'SENT';
	}
	
	private static function UpdateSendStage () {
		if (!self::$SentAsInfo && !(self::$SentAsInfo= self::$SQ->QueryStageSendAs(self::$StageId)))
			return FALSE;

		$RecipientId= is_bigint(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientId]) ? self::$aStage[CheckCaseHolderForStage_StageBase + Stage_RecipientId] : 0;
		$bAddShowToClient= FALSE;
		if ($RecipientId) {
			if (!$aPossibleViewers= self::$SQ->QueryCaseClientsAndContacts(self::$aStage[CheckCaseHolderForStage_StageBase + Stage_CaseID]))
				return FALSE;
			foreach ($aPossibleViewers as $ViewerInfo)
				if ($ViewerInfo[0] == $RecipientId) {
					$bAddShowToClient= TRUE;
					$NotifiedAt= $_SESSION['ProcessngStageEmailType'] ? MySQL_NowValue : NULL;
					$GrantAccessSQL= "INSERT INTO `tblcasestageaccess` (`EntityID`, `StageID`, `NotifiedAt`) VALUES (?, ?, ?) ON DUPLICATE KEY UPDATE `NotifiedAt`= ?";
					$ParamArray= array($RecipientId, self::$StageId, $NotifiedAt, $NotifiedAt);
					$aParamsTypes= array (MySQL_int, MySQL_int, MySQL_date, MySQL_date);
					if (!self::$SQ->SQL_CheckRoleBack($GrantAccessSQL, $Error, $ParamArray, $aParamsTypes))
						return FALSE;
					break;
				}
		}

		if ($bAddShowToClient) {
			$ReportSQL= "UPDATE `tblcasestage` SET  `ReportToClient`= ? WHERE `CaseStageID` = ?";
			$ParamArray= array (1, self::$StageId);
			$aParamsTypes= array (MySQL_int, MySQL_int);
			return self::$SQ->SQL_CheckRoleBack($ReportSQL, $Error, $ParamArray, $aParamsTypes);
		} else
			return TRUE;
	}
}
