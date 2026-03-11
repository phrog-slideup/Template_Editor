<?php 

  date_default_timezone_set('UTC');  // ensure UTC timezone
  $serverDate = date('Y-m-d');
  // echo $serverDate;
  helper('commonuserdetails');
  $companyId = (int) getUserCompanyid();
  $session = \Config\Services::session();
  $session_data = $session->get();

  $cloginuserid = getUserId();
  $issadminuser = checkAdmin($cloginuserid);

  if (!isset($session_data['user']['username'])) {
    header('Location: ' . site_url('login'));
    exit;
  } 
  $umail= $session_data['user']['username'];
  $builder = \Config\Database::connect()->table('deckusers');
  $userData = $builder->select('first_name')->where('username', $umail)->get()->getRow();

  $domainExist=true;


  if (isset($_GET['searchkey'])) {
    $searchkey = $_GET['searchkey'];
  } else {
    $searchkey = '';
  }
  $domaincatfilter = json_decode(json_encode($producttag));
  // $is_slideuplift =0;
  if (!$is_slideuplift || $is_slideuplift !== '1') : ?>
    <style>
      .dropdown-contentcat {
        top: 75px !important;
        left: 0px;
      }
    </style>
  <?php endif; ?>


<html>
  <head>
    <style>
       .chat-area{
        padding: 0px 1px !important;
      }
      .chat-content{
        height:auto !important;
      }
      .main-content{
        margin-bottom:0px;
      }
      .chat-section {
        overflow:clip !important;
      }
      @media (min-width: 1231px) and (max-width: 1400px) {
        .chat-area {
            height: 100%!important;
        }
      }
    </style>
    <!-- Preload critical resources -->
    <link rel="preload" href="<?= base_url('assets/js/unifieddeckslide.js') ?>" as="script">
    <link rel="preload" href="<?= base_url('assets/css/deckeditor.css') ?>" as="style">
    <link rel="preload" href="<?= base_url('assets/js/editor.js') ?>" as="script">
    <!-- <link rel="preload" href="<?// base_url('assets/js/aiagents.js') ?>" as="script"> -->
    <link rel="preload" href="<?= base_url('assets/js/deckagent.js') ?>" as="script">

    <?= $this->include('layouts/header') ?>
    <!-- Load CSS files -->
    <link rel="stylesheet" href="<?= base_url('assets/css/deckeditor.css') ?>">
    <link rel="stylesheet" href="<?= base_url('assets/css/editor.css') ?>">
    <link rel="stylesheet" href="<?= base_url('assets/css/agenda.css') ?>">
    <link rel="stylesheet" href="<?= base_url('assets/css/unifiedaiagent.css') ?>">
     <link rel="stylesheet" href="<?= base_url('assets/css/editor-styles.css') ?>">

    <link rel="preload" href="https://cdn.jsdelivr.net/npm/owl.carousel/dist/assets/owl.carousel.min.css" as="style" onload="this.onload=null;this.rel='stylesheet'">
    <link rel="preload" href="https://cdn.jsdelivr.net/npm/owl.carousel/dist/assets/owl.theme.default.min.css" as="style" onload="this.onload=null;this.rel='stylesheet'">
    
    <?php
      if ((int)$companyId === 3) {
        // For company 3 → true only if member
        $showDownload = !empty($membershipDetails['isMembership']);
      } else {
        // All other companies → always true
        $showDownload = true;
      }
    ?>
    <script>
      var showDownloadLink = <?php echo $showDownload ? 'true' : 'false'; ?>;
      //  console.log("sad", showDownloadLink);
    </script>

    <script>
      //JS variables
      const currentdeckid = "<?php echo (isset($_GET['deckid'])) ? $_GET['deckid'] : 0 ; ?>";
      var noofslidecount;
      var isInsertingSlides = false;
      var globalSlideIds = [];
      var userloinemail ='<?php echo $umail; ?>';
      var hostedurl = '<?php echo hostedurl; ?>';
      var isAdminUser = <?php echo json_encode($issadminuser); ?>;
      var user_login_id=<?php echo json_encode($cloginuserid); ?>;
      const currentUtcDate = "<?php echo $serverDate; ?>";
      var deckoffset = 0; 
      var decklimit = 12; 
      var searchStringText='';
      var controllerdeck;
      var currentdeckRequest=null;

      var fileedit;
      var updatelielement;
      var originalslideid;
      var idtoupdateinarray;//if there are same template then which index template to
      var opendoctype;
      var timer;

      let progresscircle = document.getElementById("progress");
      let timerText = document.getElementById("timerText");
      let timeLeft = 30;
      let interval;
      var pptdeckoffset =0;
      var pptdecklimit =12;
      var dfoffset = 0; 
      var dflimit = 12; 
      let timeoutId;
      var searchStringText='';
      var soffset = 0; 
      var slimit = 12;
      var searchids;
      //single_page search start
      var slideoffset = 0; 
      var slidelimit = 12; 
      var currentIndex = 0;
      var searchdebounceTimeout;
      let timeoutIdmore;
      let isLoading = false;  
      let currentRequest = null;
      var offset = 0; 
      var limit = 12; 
      // Define dragInBetween globally as in your original code
      var dragInBetween = -1;
      var onedriveslideid;
      var progress = 100;
      var onedriveslidesize;
      var duration;
      var mainNodePpt='';
      var recentlyClickedItem=0;
      var pptdata = null;
      var syncingIntervals = [];
      var templateCounter = 0;
      var templateCounterInsertall = -1;
      var imgPaths = [];
      var userSelectedDecks = [];
      var globalDeckid=0;
      var isSyncingStart = false;
      var deckImages = [];
      // JS variables
      var clickedType = '';
      var slidesOffset = 0;
      var decksOffset = 0;
      var slideScrollTop=0;
      var deckScrollTop=0;
      var dragInBetween = -1;
      var dragImageSrc = '';
      var globalCountdown = 0;
      var editLinkObj = ''; 
      var liElementId = ''; 
      var globalClickedType=0;
      var WINDOWS_SERVER_IP = "https://<?=WINDOWS_SERVER_IP?>/api/";
      var APPRAISAL_DECK_IDS = <?=json_encode(APPRAISAL_DECK_IDS)?>;
      var WINDOWS_SERVER_DOWNLOAD_IP = "https://<?=WINDOWS_SERVER_IP?>/outputfiles";
      const home_url = '<?=home_url()?>';
      const RAJDARBAR_DECK_ID = <?=RAJDARBAR_DECK_ID?>
      
     var sul_copy_deck = false;
      <?php if($session->get('SUL_COPY_DECK')) { ?>
         sul_copy_deck = <?=$session->get('SUL_COPY_DECK')?>; 
      <?php $session->remove('SUL_COPY_DECK'); } ?>


      var isPartnerValue = false;

      window.addEventListener('DOMContentLoaded', () => {
        const chatState = localStorage.getItem('listChatState');
        if (chatState !== null && chatState !== '') {
          agendacreation(); 
        }
      });

      <?php 
      if(isset($is_partner) && $is_partner == 1) { ?>
        isPartnerValue = true;
      <?php } ?>

      var partnerCompanyId = '<?=PARTNER_COMPANY_ID?>';
      var mohawkindCompanyId = <?= (int)MOHAWKIND_COMPANY_ID?>; 
      var isNeoPopup = false;

      <?php if($session->get('agentopensec')) { ?>
        isNeoPopup = true;
      <?php } ?>
      
      var isMembership = false;

      <?php  if (isset($membershipDetails) && !isset($membershipDetails['membershipId'])) { ?>
        isMembership = true;
      <?php } ?>

      var getCompanyTable = '<?= getCompanyTable()?>';
      const getCompanyTableActual = '<?= getCompanyTable(true)?>';

      var goldUser = '<?=PARTNER_COMPANY_NAME?>';
      const base_url = '<?=base_url()?>';
      var is_deck_id = false;

      <?php   if (isset($_GET['deckid'])) { ?>
        is_deck_id = true;
      <?php } ?>

      var company_id_value = <?php echo $companyId; ?>;

      var is_templateId = false;
      var filteroffset = 0; 
      var filterlimit = 12;
      var previousselection =[];

      const isPartnerExist = <?php echo isset($is_partner) && $is_partner === "1" ? 'true' : 'false'; ?>;
      const hasDeckId = <?php echo isset($_GET['deckid']) ? 'true' : 'false'; ?>;
      const hasDeckIdBool = <?php echo isset($_GET['deckid']) ? $_GET['deckid'] : 0; ?>;
      const user_email = '<?php echo $email;?>';
      const tokenValue = '<?php echo isset($_COOKIE["apitoken"]) ? $_COOKIE["apitoken"] : ""; ?>';
      var searchKeyVal = '<?=$searchkey?>';

      <?php if (isset($_GET['templateId'])) { ?>
        is_templateId = <?=$_GET['templateId']?>;
      <?php } ?>

      var editorId = <?= (int) isset($_GET['deckid']) ? $_GET['deckid'] : 0 ?>;

      window.LAST_INSERTED_SLIDE_IDS = [];



      $(document).ready(function() {

        // Listen for the custom event

window.addEventListener('slideIdsUpdated', async function(e) {

console.log('slideIdsUpdated event listener called');

const templateIds = e.detail.ids; // array of ids

// Call your function here

for (const id of templateIds) {

console.log('react should call this function on change ',id);
//return new Promise((resolve) => {
 await neo_insert_template(id, tablename);
//});
// your logic here

}

});

function updateLiBgImageTimestamp(liEl) {

if (!liEl) return;

// Get current background-image value

const bgImage = liEl.style.backgroundImage;

if (!bgImage) return;

// Extract URL from `url("...")`

const urlMatch = bgImage.match(/url\(["']?(.*?)["']?\)/);

if (!urlMatch) return;

let url = urlMatch[1];

const timestamp = Date.now();

// Remove existing query string (if any)

url = url.split('?')[0];

// Add fresh timestamp

const newUrl = `${url}?t=${timestamp}`;

// Set updated background-image

liEl.style.backgroundImage = `url("${newUrl}")`;

}

function indexOfId(arr, id) {

const target = String(id);

return arr.findIndex(x => String(x) === target);

}

// Add event listener

window.addEventListener('themeModal', (event) => {
  var slideIds = event.detail.slideIds;
  npvShowGallery(slideIds);
});

// Add event listener
window.addEventListener('slidesUpdated', (event) => {
var slideId = event.detail.slideIds;
console.log('Slides updated:', slideId);

const li = document.querySelector('li[slide_id="'+ slideId +'"]');
updateLiBgImageTimestamp(li);
});

window.USER_CONTEXT = {
  email: user_email,
  userId: user_login_id,
  companyId: company_id_value,
  deck_id: editorId
}
        <?php if(isset($_GET['admindeckid'])) { ?>
          admindeckid = <?php echo $_GET['admindeckid']; ?>;
          insertallslide(admindeckid);
      
        <?php } ?>

        <?php if(isset($_GET['othersdecksid'])) { ?>
          othersdecksid = <?php echo $_GET['othersdecksid']; ?>;
          insertallslide(othersdecksid);
      
        <?php } ?>

        <?php 
          if (!isset($_GET['deckid'])) { 
            if ($session->get('agentopensec') === 'showaagentdiv'){ ?>
              <?php if (!isset($session_data['slideId'])) {?>
                createnewdeckModal('aiagent');
              <?php } else { ?>
                createnewdeckModal('aiagent');
              <?php }?>
            <?php
            }
            else{ ?>
              createnewdeckModal();
        <?php } } ?>

        <?php if (isset($_GET['deckid']) && isset($session_data['slideId'])) { 
            $slideIds = $session_data['slideId'];
            if (!is_array($slideIds)) { 
                $slideIds = [$slideIds];
            }
          ?>
          const templateIds = <?= json_encode($slideIds) ?>;

          (async () => {
              for (const id of templateIds) {
                  console.log('react should call this function on change ', id);
                  await neo_insert_template(id, tablename);
              }
          })();

          <?php $session->remove('slideId');  
        } ?>

        <?php if (isset($_GET['aiagent']) && $_GET['aiagent'] == '1') { ?>
          var url = new URL(window.location);
          url.searchParams.delete('aiagent');
          window.history.replaceState({}, '', url); 
          removeloadedscript();
          agendacreation(); 
        <?php } ?>

        <?php if(isset($_GET['copyslideid'])) { ?>
          var erptablename="<?php echo getCompanytable(); ?>";
          copyslideid = <?php echo $_GET['copyslideid']; ?>;
          insert_template(copyslideid,erptablename)
          var url = new URL(window.location);
          url.searchParams.delete('copyslideid');
          window.history.replaceState({}, '', url);
        <?php } ?>

      });
    </script>
    <title>Neo - AI Presentation Maker</title>
  </head>


  <body>
    <div id="spinnerOverlaybase" class="inlineslide-loading" style="display: none;">
      <!-- Your spinner content -->
    </div>
    <script>
      document.getElementById('spinnerOverlaybase').style.display = 'flex';
      setTimeout(() => {
        const spinnerOverlaybase = document.getElementById('spinnerOverlaybase');
        if (spinnerOverlaybase) {
          spinnerOverlaybase.style.display = 'none';
        }
      }, 5000);
    </script>
    <!-- Modal Carousel -->
    <div class="deckpreviewmodal" id="previewModal">
      <div class="popup-overlay" onclick="closedeckpreviewPopup()"></div>
      <div class="carousel-container">
        <a onclick="closedeckpreviewPopup()" class="closeslidepreview">
          <img src="<?= base_url('images/Vector2.svg') ?>" style="width:12px;height:12px">
        </a>
        
        <div class="carousel-track">
            <!-- Slides will be dynamically inserted here -->
        </div>
        
        <div class="carousel-nav">
          <button class="carousel-prev">
            <img src="<?= base_url('images/slidearrow.svg') ?>" width="16" height="16"  >
          </button>
          <div class="numberindicator">
            <div class="carousel-indicators">
              <!-- Indicators will be dynamically inserted here -->
            </div>
            <p id="carousel-counter" class="carousel-counter">1/1</p>
          </div>
          <button class="carousel-next">
            <img src="<?= base_url('images/slidearrow.svg') ?>" width="16" height="16" style="transform: rotate(180deg);" >
          </button>
        </div>
      </div>
    </div>

    <script>
      var globalFilename='';
    </script>

    <?php if(isset($_GET['deckids'])) { ?>
      <script>
        var globalFilename='<?php echo basename($deck_details->downloadurlpath);?>';
      </script>

      <div id="overlays" class="overlays">
        <div id="loader" class="loader"></div>
      </div> 
    <?php  }?> 

    <div id="ai-popupModal" class="ai-popup-modal">
      <div class="ai-popup-content">
        <a onclick="closePopupai()" class="cancelai"><img src="<?= base_url('images/Vector2.svg') ?>"  alt="Image" /></a>

        <form action="" class="requestform">
          <div class="aiinput" id="myaiinput">
            <textarea id="airequest" name="airequest" placeholder="Redo slide with AI" oninput="slidelevelinput(this)" onkeydown="checkEnterKeyforslideai(event)"></textarea>
            <img onclick="updatewithai()" src="<?= base_url('images/upwardarrow.svg') ?>" style="width: 24px;" class="disabled" id="submitImage">
          </div>
          <a class="aibtn" onclick="slideaiinputtoggle(this)">
            <img class="regenerateai aibtnregenerateai" src="<?= base_url('images/ai.svg') ?>"  style="width: 14px;">
          </a>
        </form>

        <div class="aipopupmain">
          <div class="ai-popup-image" style="position: relative;">
            <img id="ai-popupImage" src="" alt="Image" />
            <div class="manualtextimgloader" id="loader-overlay" style="display: none; position: absolute; top: 0; left: 0; width: 100%; height: 100%; background-color: rgba(0, 0, 0, 0.5);justify-content: center; align-items: center;">
              <img src="<?= base_url('images/loadinganimiwhite.gif') ?>" alt="Loading..." style="width: 50px; height: 50px;">
            </div>
          </div>
          <div class="ai-popup-data">
            <div  id="aitextbox">
              <div class="hdngsave" id="titlediv">
                <img src="<?= base_url('images/ograpics.svg') ?>" alt="Loading..." class="titleedit" onclick="showaititle(event)">
                <h1 class="aiheading" id="titleppt"><input type="text" value="" id="sullytitle" name="sullytextplace" ></h1>

                <div class="textlevelai">
                  <a class="generatetext" onclick="toggleInputVisibility(this)">
                    <img class="regenerateai" src="<?= base_url('images/ai.svg') ?>" style="width: 14px;">
                  </a>
                  <div class="textlevelinput">
                    <input type="text" id="sullytitleinput" name="sullytitle" placeholder="What can I help with ?" oninput="checkInput(this)" >
                    <img onclick="updatetextboxesai('sullytitle',event)" src="<?= base_url('images/upwardarrow.svg') ?>" style="width: 22px;" class="disabled">
                  </div>
                </div>
              </div>
            </div>
            <div class="savecancel">
              <a onclick="savemanualtext()" class="savetext"><img src="<?= base_url('images/saveai.svg') ?>" alt="Image" />Save</a>
            </div>
          </div>
        </div>
      </div>
    </div>

    <div id="videoPopup" class="overlayvdo">
      <div class="vdopopup-content">
        <span id="closePopupBtn" class="close-btn">&times;</span>
        <video id="video" controls>
            <source src="<?= base_url('images/product_demo.mp4') ?>" type="video/mp4">
            Your browser does not support the video tag.
        </video>
      </div>
    </div>

    <div class="savepopup-overlay">
      <div class="savepopup">
        <svg class="savepopup-icon" viewBox="0 0 24 24">
          <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
        </svg>
        <div class="savepopup-content">
          <h2 class="savepopup-title">Deck saved</h2>
          <p class="savepopup-message">All changes have been saved to drafts</p>
        </div>
      </div>
    </div>

    <div id="overlay1" class="overlay1" style="display: none;">
      <div class="modal1">
        <span class="close" id="pnamesavepopup" onclick="closesaveblock()" ><img src="<?= base_url('images/Vector2.svg') ?>" ></span>
        <p class="modal-text1">Give your presentation a name</p>
        <input id="prodeckname" type="text" class="edittextheader" placeholder="Project Name" value="Untitled Presentation">
        <div class="button-group1">
            <button onclick="validateProjectName(true)" class="action-btn1" id='btn_save_project_name' >Save</button>
            <button onclick="hardReload()" class="action-btn1" id="discard" >Discard</button>
        </div>
      </div>
    </div>

    <script>
      var statStatus = false;
    </script>

    <div class="section">
      <div class="container-fluid maincolr">
        <div class="bannner-type"> 
          <?php 
            if (!isset($isEditor)) {
              $isEditor1 = false;

              $base_url = base_url();
              $deckeditorurlorg1 = $base_url . 'deckeditor';

              $currenturlorg1 = "http://" . $_SERVER['HTTP_HOST'] . $_SERVER['REQUEST_URI'];
              $current_url1 = parse_url($currenturlorg1, PHP_URL_PATH);
              $deckeditorurl1 = parse_url($deckeditorurlorg1, PHP_URL_PATH);
              if ($current_url1 == $deckeditorurl1) {
                $isEditor1 = true;
              } 
            }
          ?>

          <!-- Agent Modal -->
          <div class="agent-modal-overlay" id="agentModalOverlay">
            <div class="agent-modal">
              <div class="agent-modal-header">
                <h2 class="agent-modal-title">Choose Slide Type</h2>
                <p class="agent-modal-subtitle">Select from 133+ professional slide templates</p>
                <button class="agent-modal-close-button" id="agentCloseButton"><img src="<?= base_url('images/Vector2.svg') ?>" ></button>
              </div>

              <div class="agent-modal-content">
                <div class="slide-categories">
                  <div class="agentcategory">
                    <div class="agentcategory-header">
                      <h3 class="agentcategory-title">Introduction</h3>
                    </div>
                    <ul class="slide-list">
                      <li class="slide-item" onclick="handleTopicButtonClick('agenda')">Agenda Slide</li>
                      <li class="slide-item" onclick="handleTopicButtonClick('title')">Title Slide</li>
                      <li class="slide-item">SCQA Slide</li>
                      <li class="slide-item">Introduction Slide</li>
                      <li class="slide-item">Problem Statement Slide</li>
                      <li class="slide-item hidden">Objective Slide</li>
                      <li class="slide-item hidden">Icebreaker Slide</li>
                    </ul>
                    <div class="show-toggle" data-category="introduction">Show more...</div>
                  </div>

                  <div class="agentcategory">
                    <div class="agentcategory-header">
                      <h3 class="agentcategory-title">Conclusion</h3>
                    </div>
                    <ul class="slide-list">
                      <li class="slide-item" onclick="handleTopicButtonClick('thanks')">Thank You Slide</li>
                      <li class="slide-item">Conclusion Slide</li>
                      <li class="slide-item">Summary Slide</li>
                      <li class="slide-item">Call to Action Slide</li>
                      <li class="slide-item">Q&A Slide</li>
                      <li class="slide-item hidden">Contact Information Slide</li>
                      <li class="slide-item hidden">Resources Slide</li>
                    </ul>
                    <div class="show-toggle" data-category="conclusion">Show more...</div>
                  </div>

                  <div class="agentcategory">
                    <div class="agentcategory-header">
                      <h3 class="agentcategory-title">Project Planning</h3>
                    </div>
                    <ul class="slide-list">
                      <li class="slide-item" onclick="handleTopicButtonClick('timeline')">Timeline Slide</li>
                      <li class="slide-item" onclick="handleTopicButtonClick('roadmap')">Roadmap Slide</li>
                      <li class="slide-item">Gantt Chart Slide</li>
                      <li class="slide-item">Process Slide</li>
                      <li class="slide-item">Next Steps Slide</li>
                      <li class="slide-item hidden">Risk Assessment Slide</li>
                      <li class="slide-item hidden">Milestone Slide</li>
                    </ul>
                    <div class="show-toggle" data-category="project">Show more...</div>
                  </div>

                  <div class="agentcategory">
                    <div class="agentcategory-header">
                      <h3 class="agentcategory-title">Strategic Analysis</h3>
                    </div>
                    <ul class="slide-list">
                      <li class="slide-item hidden">Problem Solution Slide</li>
                      <li class="slide-item">SWOT Analysis Slide</li>
                      <li class="slide-item">PESTLE Analysis Slide</li>
                      <li class="slide-item">BCG Matrix Slide</li>
                      <li class="slide-item">McKinsey 7S Slide</li>
                      <li class="slide-item">Business Model Canvas Slide</li>
                      <li class="slide-item hidden">Competitor Analysis Slide</li>
                      <li class="slide-item hidden">Market Analysis Slide</li>
                      <li class="slide-item hidden">Comparison Slide</li>
                      <li class="slide-item hidden">Concept Slide</li>
                    </ul>
                    <div class="show-toggle" data-category="strategic">Show more...</div>
                  </div>

                  <div class="agentcategory">
                    <div class="agentcategory-header">
                      <h3 class="agentcategory-title">Financial Data</h3>
                    </div>
                    <ul class="slide-list">
                      <li class="slide-item" onclick="handleTopicButtonClick('graph')">Chart Slide</li>
                      <li class="slide-item">Financial Projections Slide</li>
                      <li class="slide-item">Budget Slide</li>
                      <li class="slide-item">Table Slide</li>
                      <li class="slide-item">Statistics Slide</li>
                    </ul>
                  </div>

                  <div class="agentcategory">
                    <div class="agentcategory-header">
                      <h3 class="agentcategory-title">Company Profile</h3>
                    </div>
                    <ul class="slide-list">
                      <li class="slide-item"  onclick="handleTopicButtonClick('fandb')">Feature Benefits Slide</li>
                      <li class="slide-item"  onclick="handleTopicButtonClick('testimonial')">Testimonial Slide</li>
                      <li class="slide-item"  onclick="handleTopicButtonClick('pandc')">Pros & Cons Slide</li>
                      <li class="slide-item">Team Slide</li>
                      <li class="slide-item">Mission Vision Values Slide</li>
                      <li class="slide-item">Product Overview Slide</li>
                      <li class="slide-item hidden">Disclaimer Slide</li>
                      <li class="slide-item hidden">Copyright Slide</li>
                    </ul>
                    <div class="show-toggle" data-category="company">Show more...</div>
                  </div>

                  <div class="agentcategory">
                    <div class="agentcategory-header">
                      <h3 class="agentcategory-title">Visual Elements</h3>
                    </div>
                    <ul class="slide-list">
                      <li class="slide-item">Collage Slide</li>
                      <li class="slide-item">Wordcloud Slide</li>
                      <li class="slide-item">Image Slide</li>
                      <li class="slide-item">Infographic Slide</li>
                      <li class="slide-item">Video Slide</li>
                    </ul>
                  </div>

                  <div class="agentcategory">
                    <div class="agentcategory-header">
                      <h3 class="agentcategory-title">Core Information</h3>
                    </div>
                    <ul class="slide-list">
                      <li class="slide-item">Text Slide</li>
                      <li class="slide-item">Key Points Slide</li>
                      <li class="slide-item">Quote Slide</li>
                      <li class="slide-item">Definition Slide</li>
                      <li class="slide-item">Case Study Slide</li>
                    </ul>
                  </div>

                  
                  <div class="agentcategory">
                    <div class="agentcategory-header">
                      <h3 class="agentcategory-title">Engagement</h3>
                    </div>
                    <ul class="slide-list">
                      <li class="slide-item" onclick="handleTopicButtonClick('break')">Break Slide</li>
                      <li class="slide-item">Quiz Slide</li>
                      <li class="slide-item">Question Slide</li>
                      <li class="slide-item">Interactive Poll Slide</li>
                      <li class="slide-item">Discussion Slide</li>
                    </ul>
                  </div>
                </div>
              </div>
            </div>
          </div>
          <!-- Agent Modal end-->

          <div class="profile  <?= ($isEditor1 == true) ? 'collapsed':''?>" id="profile">
            <?php echo view('layouts/sidebarcomponent'); ?>  
          </div>
      
          <div class='chat-section panel-open <?= ($isEditor1 == true) ? 'expand':''?>' id="chatSection">

            <?php if($companyId == 3){ ?>
              <a onclick="openneopopup()" class="freeneo"><img class="" src="<?= base_url('images/neo.gif') ?>">NEO</a>
              <div class="neo-popup">
                <div class="neo-header">
                  <div class="neo-title">
                    <span><img style="margin-right: 0.5rem;" src="<?= base_url('images/ai2.svg') ?>" width="24" alt="">NEO</span>
                    <span class="neo-status"><span class="status-dot"></span>Ready</span>
                  </div>
                  <div class="neo-actions">
                    <span class="neo-close" onclick="closeneopopup()"><img style="filter: brightness(0.5);" width="12px" src="<?= base_url('images/Vector2.svg') ?>" ></span>
                  </div>
                </div>
                <div class="neo-body">
                  <p>Hi! I'm NEO, your AI presentation assistant. I can help you create professional slides quickly.</p>
                  <div class="neo-help">
                    <p>What I can help with:</p>
                    <ul>
                      <li>Agendas</li>
                      <li>Testimonials</li>
                      <li>Timelines</li>
                      <li>Features</li>
                    </ul>
                  </div>
                </div>
                <button class="neo-button" onclick="startneo()">Start Creating</button>
              </div>
            <?php } ?>


            <?php echo view('layouts/headercomponent'); ?>

            <?php
              $isSulReferral = $session->get('agentopensec') ?? '';
              if (is_null($deck_details) && $isSulReferral == '') { 
                header('Location: ' . url_to('workspace'));  
                exit; // Stop further execution of PHP
              }
            ?>

            <?php if($companyId == 3){ ?>
              <?php echo view('layouts/feedbackform'); ?>
            <?php } ?>

            <div id="filterOverlay" class="filter-overlay" ></div>
            <div class='chat-content'>
              <div class="maincontent">
                <div class="hed_div" id="pro_00022" style="display:none">
                  <input type="text" id="edittextheader1" placeholder="Project Name" style="<?php if(!isset($_GET['deckid'])) { echo 'display:none;';}?>" value='<?php if(isset($_GET['deckid'])) { echo $deck_details->projectname;} else { echo "Untitled Presentation";} ?>'>
                  <div id="message" style="color: #d53939;margin-top: 4px;"></div>            
                </div>

                <div class="actionbutton" id="actionbutton" style="<?php if(!isset($_GET['deckid'])) { echo 'display:none;';}?>"> 
                  <div class="loadersync">
                    <p style="margin: 0;color:#6467a9">Syncing</p>
                    <div class="loadersyncing"></div>
                  </div>

                  <a style="display:none" class="downppts mobilenone" onclick="sync(this,syncprev=true)" id='idsyncbtn' style="font-weight: bold;  cursor: pointer;color: #6467a9;
                    font-size: 13px;">
                    <span class="dnldicon1 forremove" id="saveiconfaicons">Changes not showing up? </span>
                    <span class="forremove" style="text-decoration: underline;">Sync Now</span>  
                  </a>
              
                  <a style="display:none" class="downppts mobilenone hide" id='syncgif' style="font-weight: bold;cursor: pointer;color: #6467a9;
                    font-size: 13px;">
                    <span class="dnldicon1 forremove" id="saveiconfaicons">Changes not showing up? </span>
                    <span class="forremove" style="text-decoration: underline;">Sync Now</span>
                  </a>

                  <div class="progress-wrapper">
                    <p class="timer-text" id="timerText-primary">Starting sync engine...</p>
                    <p class="timer-text hide" id="timerText">Next sync within 30s...</p>
                    <div class="progress-container">
                      <div class="progress" id="progress"></div>
                    </div>
                  </div>

                  <?php if($domainExist){ 
                    $clsOneDrive = '';
                    if ($merged_presentation == null || $merged_presentation == '' || $merged_presentation == 0) {
                      $clsOneDrive = ' hide ';
                    } ?>
                
                    <!-- <a id="btnSyncOneDrive" class=" mobilenone<?= $clsOneDrive?>" onclick="syncDeckChanges(this)">Sync with OneDrive</a> -->
                    <p id="timerMessage">Ready to sync in...<span id="countdown">45</span> sec</p>

                    <a id="save_download" class="downppt mobilenone forcehiddenelement" onclick="downloadDeck(true)">
                      <img class="dnldicon" id="saveiconfaicon" src="<?= base_url('images/save-instagram.png') ?>">
                      <img src="<?= base_url('images/loadinganimiwhite.gif') ?>"  width="18" height="18" id="downloadloader" style="display:none;">Save     
                      <span class="tooltip">Click to Save the Deck</span>
                    </a>
             
                    <!-- ddl code started  -->
                    <div class="dropdown forcehiddenelement">
                      <a id="onlydownload" class="downppt mobilenone forcehiddenelement" onclick="toggleDropdown(event)">
                        <img class="dnldicon" id="downloadfaicon" src=<?= base_url('images/exportppt.svg') ?> style="">
                        <img src="<?= base_url('images/loadinganimiwhite.gif') ?>" width="18" height="18" id="downloadloader1" style="display: none;">Export
                      </a>
                  
                      <div id="dropdownMenu" class="dropdown-content" style="display: none;">
                        <a href="javascript:void(0)" onclick="downloadDeck(); hideDropdown()"><img class="" style="filter:unset;width:18px" id="downloadfaicon" src="<?= base_url('images/download-icon-vq.svg') ?>">Download</a>
                        <a href="javascript:void(0)" onclick="downloadDeck(false, 'onedrive'); hideDropdown()"><img class="" style="filter:unset;width:18px" id="downloadfaicon" src="<?= base_url('images/onedrive.svg') ?>">Open in OneDrive</a>
                      </div>
                    </div>
                    <!-- ddl code end -->
                  <?php } ?>
              
                  <a class="contentMenu downppt" id="contentMenuId" onclick="showActionlink()" style="display: none;"> 
                    <i class="fa fa-bars" aria-hidden="true" id="menubarDeck"  style="display:block;"></i>
                    <i class="fa fa-times" aria-hidden="true" id="crossbarDeck"  style="display:none;"></i>
                  </a>

                  <div class="deckAction" id="deckActionId" style="display:none;">
                    <?php if($domainExist) { ?>
                      <a id="save_download" class="actionelink" onclick="downloadDeck()" >
                        <i class="fa fa-download" aria-hidden="true"></i>
                        <span class="texticon"> Save & Download</span>     
                      </a>
                    <?php } ?>

                    <a href="javascript:void(0)" class="actionelink" id="send_service" onclick="sendService()" >
                      <i class="fa fa-cogs" aria-hidden="true" id="serviceicon"></i> 
                      <img src="<?= base_url('images/deckbuilder.gif') ?>"  width="30" height="30" id="serviceloader" style="display:none;">
                      <span class="texticon"> Send for Service</span>
                    </a>

                    <a href="" class="actionelink" id="create_again"  onclick="location.reload(); return false;">
                      <i class="fa fa-file" aria-hidden="true"></i>
                      <span class="texticon">Create New</span>
                    </a>
                  </div>
                </div>
              </div>

              <?php if(!isset($_GET['deckid']) && !isset($_GET['newdeck'])) { ?>
                <div class="welcomemsg" id="idwelcomemsg">
                  <img src="<?= base_url('images/avtarlogo1.svg') ?>" width="150px"> 
                  <h1>Welcome <span class="deckstar"><?php echo isset($userData->first_name) ? $userData->first_name : 'D'; ?></span></h1>
                  <p>Make a deck from a curated library of slides</p>
                </div>
            
                <div class="pickfromlibrary">
                  <button class="addslidebtnmid" id="addslidebtn" onclick="createNewdeck()">
                    <img src="<?= base_url('images/library.png') ?>" id="library"><p class="blibrary">Browse Library</p>
                  </button>
                </div>
              <?php } ?>

              <div id="popupBackdrop" class="backdrop" style="display: none;">
                <div class="inputtextpopup">
                  <button id="itextclosePopup" onclick="inputtextclosePopup()"><i class="fa fa-times" aria-hidden="true"></i></button>
                  <div class="pseudo-search">
                    <button class="fa fa-search contentserachbtn" type="submit"></button>
                    <input type="text" id="fillTopic" placeholder="Search..." autofocus required>
                    <a onclick="processfiller()" style="cursor: pointer;">
                      <img src="" style="max-width: 20px;">
                    </a>
                  </div>
                </div>
              </div>

              <div id="editmodalBackdrop" class="backdrop" style="display: none;">
                <div id="editorModal" class="modal">
                  <div id="saveupdate" >
                    <img src="<?= base_url('images/edittextloader.gif') ?>" style="max-width: 40px;">
                  </div>
                  <div class="modal-content edittext_modalcontent">
                    <div class="editmodal_action">
                      <input type='button' value='Save' class='' id='btnsaveEditor' />
                      <span class="close" id="closeframe" onclick="closeiframe()">&times;</span>
                    </div>

                    <div class='col-md-12' id='loading-gif' style='top:50%;text-align: center;flex:0;'></div>
                    <div class="edittext_div col-12">
                      <div class='col-md-6' id='images'></div>
                      <div class='col-md-6' id='manualedit'></div>
                    </div>
                  </div>
                </div>
              </div>

              <?php 
                $session = \Config\Services::session();
              ?>

              <div id="slideContainer" class="visible">                
                <div id='filterhook'></div>             
                  <div id="filterOverlay2" class="filter-overlay" ></div>
                  <?php if($companyId != 4){ ?>
                  <div class="left-div" id="agendaview" >
                    <div class="chat-area" id="agendaarea">
                      <div id="ratelimitexceed">
                        <div class="ratelimitimgsec">
                          <img class="waringicon" id="" src="<?= base_url('images/warninglimit.svg') ?>" >
                          <img class="closewarning" onclick="closelimitwarning()" src="<?= base_url('images/Vector2.svg') ?>" >
                        </div>
                        <p class="ratelimttext" >Daily AI usage limit reached. Please come back tomorrow.</p>
                      </div>

                      <div id="filesizelimitexceed">
                        <div class="ratelimitimgsec">
                          <img class="waringicon" id="" src="<?= base_url('images/warninglimit.svg') ?>" >
                          <img class="closewarning" onclick="closelimitwarning()" src="<?= base_url('images/Vector2.svg') ?>" >
                        </div>
                        <p class="filesizelimiterror" >File Size must be less than 10MB</p>
                      </div>

                      <div class="initialmessage" id="initialmessagediv" style="display:none">
                        <div class="deckslidetype" id="deckslidetypeid">
                          <p class="aiassist"> Hi, I'm Neo. I can help you create a complete deck, or build your presentation one slide at a time. How would you like to start?</p>
                        </div>

                        <div class="slidetype" id ="slidetypeid">
                          <div class="slideagent" id="slideagentid">
                            <p class="aiassist" style="margin-bottom: 1rem;"> Hi, I'm Neo. I can help you build your presentation one slide at a time. </p>
                            <?php
                              $topicbuttons = get_topic_buttons();
                            ?>
                            <?php foreach ($topicbuttons as $button): ?>
                              <button class="first-button" onclick="handleTopicButtonClick('<?php echo htmlspecialchars($button['archtype']); ?>')" data-topic-id="<?php echo htmlspecialchars($button['id']); ?>">
                                <img class="" src="<?php echo base_url('images/' . $button['imgsrc'] ); ?>" >
                                <?php echo htmlspecialchars($button['label']); ?>
                              </button>
                            <?php endforeach; ?>
                            <a class="seeall-button" id="seeAllButton">Show All</a>
                          </div>
                        </div>
                      </div>
                      <div class="startcreting" id="startcreatingdiv" style="display:none">
                        <!-- edit with neo start -->
                        <div id="aicontetentdiv" style="display:none;">
                          <div class="chatforaicontent" id="chatforaicontent">

                            <div class="messagebotinital">
                              <div class="message-labelinital">
                                <div class="message-icon ai-icon">
                                  <img src="<?php echo base_url('images/ai2.svg'); ?>" alt="">
                                </div>
                                <span>Neo</span>
                              </div>
                              <div class="message-content">
                                <span class="message-text">How can I help you with this slide?</span>
                                <div class="designimage" id="aicontentdesignimg" >
                                  <img class="design-preview-img" src="" alt="Selected Design" />
                                </div>
                              </div>
                            </div>

                            <div class="inline-input inline-input-aicontent disabledbox" id="contentgenerationtextarea">
                              <div class="inline-text">
                                <input type="file" id="editwithneoFileInput" style="display:none;" accept=".jpg,.jpeg,.png,.pdf" multiple onchange="handlecontentFileSelection(event)">
                                <textarea class="agent-auto-resize-textarea" placeholder="Describe your topic..." 
                                  onkeypress="handleaicontent(event, this)" oninput="handleaicontent(event,this)" data-question-key="topic"></textarea>
                                <div id="contentUploadedFiles" class="uploaded-files-container" style="display:none;"></div>
                                <div class="attchandenterbutton">   
                                  <button class="inline-btn enterarrow" onclick="submitInlineAnswerai(this)" disabled>
                                    <img src="<?php echo base_url('images/agendasend.svg'); ?>" alt="">
                                  </button>
                                  <div class="attachment-wrapper">
                                    <button class="attachmentbuton" type="button" onclick="document.getElementById('editwithneoFileInput').click()" >
                                      <i class="fas fa-paperclip" style="font-size:18px;color:#3c3c3c;"></i>
                                    </button>
                                    <div class="attachmenttooltip">
                                      <p class="tooltip-line1">Click to upload a file.</p>
                                      <p class="tooltip-line2">You can upload PDFs or images (JPG, PNG).</p>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                            <div id="chatsection-aicontent">
                          
                            </div>
                          </div>
                        </div>
                        <!-- edit with neo end -->
                        <div id="messages" class="messages"></div>
                        <div class="input-section" id="inputSection"></div>
                      </div>

                      <div class="startcreting" id="unfieddiv" >
                          <script type="module" src="/deckbuilder/deckster/public/assets/js/index-BHoY8oxX.js"></script>
                          <link rel="stylesheet" href="/deckbuilder/deckster/public/assets/js/index-BqK5spfE.css">
                          <div class="json_machine_deck_body">
                            <div id="root"></div>
                          </div>
                      </div>


                      <!-- <a id="feedbackopen">Neo is brand new! <span onclick="feedbkOpenModal()" style="color: #007bff;cursor: pointer;text-decoration:underline">Share</span> your experience to help us make it better.</a> -->
                    </div>
                  </div>
                  <?php } ?>

                  <div class="left-div" id="pptdataview"  <?php 
                    if($session->get('agentopensec')) {  $session->remove('agentopensec'); ?> style="display:none;" <?php } ?> <?php if($companyId == 3){ ?>style="display:none" <?php } ?> > 
                    <?php echo view('layouts/categorydesignapi'); ?>

                    <div id="slideuplifttoogle" style="display:none !important;">
                      <p id="inHouseText">In House</p>
                      <label class="toggle-container" style="margin-bottom: 0rem;">
                          <input type="checkbox" id="sul_templates">
                          <span class="toggle"></span>
                      </label>
                      <p id="partnerText">Partner</p>
                    </div>

                    <div id="projectInfo" style="display:none;">
                      <div class="projectbackarrow">
                        <button class="back-btn" onclick="goBack()"><i class="fa fa-arrow-left" aria-hidden="true"></i></button>
                        <button id="insertPagesBtn" onclick="insertmydeck()">Insert all 10 pages</button>
                        <div id="deckpreviewmnameinside">
                          <h6 id="projectName"></h6>
                        </div>
                      </div>
                    </div>

                    <div class="tabszone tabs">
                      <div class="tab active" id="tab1" onclick="showTab(1)">
                        <img class="" id="slidetabimage" src="<?= base_url('images/slides.svg') ?>" width="18px">
                        Slides
                      </div>
                      <div class="tab" id="tab2" onclick="showTab(2)">
                        <img class="inactive-tab" id="decktabimage" src="<?= base_url('images/decktab.svg') ?>" width="18px">
                        Decks
                      </div>
                      <div class="tab-indicator"></div>
                    </div>

                    <div class="icon-container"> 
                      <div class="search-container searchInput">
                        <input type="text" class="serachslide search-box" id='getSearchTemplate' placeholder="Search slide types..." >
                        <div class="resultBox" id="searchResults">

                        </div>
                        <span class="clear-icon" id="clearSearch1" onclick="clearSearchInput()">&times;</span>
                        <span class="searchslidedeckicon" id="searchslideanddeckicon" ><i class="fa fa-search" style="color: #5d5a5a; font-size:14px"></i></span>
                      </div>
                    </div>

                    <div class="filter-icon" >
                      <div class="catebreadcrumb">
                        <a id="breadcrumb-all" onclick="resettemplatefilter(event)">All</a>
                        <span id="breadcrumb-separator1" style="display: none;"> &gt; </span>
                        <a id="breadcrumb-category" style="display: none;" onclick="openCategorySubcategories()"></a>
                        <span id="breadcrumb-separator2" style="display: none;"> &gt; </span>
                        <a id="breadcrumb-subcategory" style="display: none;"></a>
                      </div>
                      <img onclick="openFilter()" src="<?= base_url('images/Setting.svg') ?>" />
                    </div>
                    
                    <div id='tab-content-scroll'>
                      <div class="tab-content" id="content1"></div>
                      <div class="tab-content" id="content2"></div>
                      <div class="tab-content" id="deckdetails"></div>
                    </div>
                  </div>
                </div>

                <!-- inline-editor -->
  <div id="editorParking" style="display:none;">

        <!-- Editor Container -->
        <div class="editor-container">

            <!-- Top Bar -->
            <div class="top-bar">
                <div class="top-bar-left">
                    <div class="app-icon">📊</div>
                    <div class="doc-title" contenteditable="true">Presentation1.pptx</div>
                </div>
                <div class="top-bar-right">
                    <button class="window-control closee" id="pptcloseModal">×</button>
                </div>
            </div>

            <!-- Menu Bar -->
            <div class="menu-bar">
                <div class="menu-item active" data-tab="home">Home</div>
                <div class="menu-item" data-tab="insert">Insert</div>
                <div class="menu-item" data-tab="draw">Shape Format</div>
                <div class="menu-item" data-tab="imageformat">Image Format</div>

                <!-- Top Undo/Redo (PowerPoint-like placement) -->
                <div class="menu-right">
                   <div class="toolbar-group">
                    <button class="tool-btn" id="undoBtnTop" title="Undo">
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                            stroke-width="2">
                            <path d="M3 7v6h6M3 13a9 9 0 1 0 3-6" />
                        </svg>
                    </button>
                    <button class="tool-btn" id="redoBtnTop" title="Redo">
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                            stroke-width="2">
                            <path d="M21 7v6h-6M21 13a9 9 0 1 1-3-6" />
                        </svg>
                    </button>
                            </div>
                    <button class="tool-btn large" id="convertToPptx" title="Save">
                        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                            stroke-width="2">
                            <path d="M12 3v12" />
                            <path d="M8 11l4 4 4-4" />
                            <path d="M4 21h16" />
                        </svg>
                        <span>Save</span>
                    </button>
                </div>
            </div>

            <!-- Toolbar Container - Changes based on active tab -->
            <div class="toolbar-container" id="textToolPanel">

                <!-- Home Tab Toolbar -->
                <div class="toolbar home-toolbar active">

                    <!-- Save/Undo/Redo Group -->
                    <div class="toolbar-group">
                        <!-- <button class="tool-btn large" id="saveChanges" title="Save & Convert to PPTX">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <path d="M19 21H5a2 2 0 01-2-2V5a2 2 0 012-2h11l5 5v11a2 2 0 01-2 2z" />
                                <polyline points="17 21 17 13 7 13 7 21" />
                                <polyline points="7 3 7 8 15 8" />
                            </svg>
                            <span>Save</span>
                        </button> -->

                        <button class="tool-btn" id="undoBtn" title="Undo">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <path d="M3 7v6h6M3 13a9 9 0 1 0 3-6.7" />
                            </svg>
                        </button>
                        <button class="tool-btn" id="redoBtn" title="Redo">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <path d="M21 7v6h-6M21 13a9 9 0 1 1-3-6.7" />
                            </svg>
                        </button>
                    </div>

                
                    <!-- Slides Group -->
                    <div class="toolbar-group" style="display: none;">
                        <button class="tool-btn large" id="addSlideBtn">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <rect x="3" y="3" width="18" height="18" rx="2" />
                                <line x1="12" y1="8" x2="12" y2="16" />
                                <line x1="8" y1="12" x2="16" y2="12" />
                            </svg>
                            <span>Add Slide</span>
                        </button>
                    </div>

                    <!-- Font Group -->
                    <div class="toolbar-group">
                        <select class="font-select" id="fontFamily">
                            <option value="Arial">Arial</option>
                            <option value="Calibri">Calibri</option>
                            <option value="Times New Roman">Times New Roman</option>
                            <option value="Georgia">Georgia</option>
                            <option value="Verdana">Verdana</option>
                            <option value="Courier New">Courier New</option>
                            <option value="Comic Sans MS">Comic Sans MS</option>
                            <option value="Impact">Impact</option>
                            <option value="Trebuchet MS">Trebuchet MS</option>
                        </select>
                        <select class="font-size" id="fontSize">
                            <option value="8">8</option>
                            <option value="10">10</option>
                            <option value="12">12</option>
                            <option value="14">14</option>
                            <option value="16">16</option>
                            <option value="18" selected>18</option>
                            <option value="20">20</option>
                            <option value="24">24</option>
                            <option value="28">28</option>
                            <option value="32">32</option>
                            <option value="36">36</option>
                            <option value="48">48</option>
                            <option value="60">60</option>
                            <option value="72">72</option>
                        </select>
                    </div>

                    <!-- Text Formatting Group -->
                    <div class="toolbar-group">
                        <button class="tool-btn" id="boldBtn" title="Bold (Ctrl+B)">
                            <strong>B</strong>
                        </button>
                        <button class="tool-btn" id="italicBtn" title="Italic (Ctrl+I)">
                            <em>I</em>
                        </button>
                        <button class="tool-btn" id="underlineBtn" title="Underline (Ctrl+U)">
                            <u>U</u>
                        </button>
                        <button class="tool-btn" id="strikeBtn" title="Strikethrough">
                            <s>S</s>
                        </button>
                    </div>

                    <!-- Color Group -->
                    <div class="toolbar-group">
                        <div class="color-picker-wrapper">
                            <input type="color" id="textColor" value="#000000" class="color-input">
                            <label for="textColor" class="color-label" title="Text Color">
                                <div class="color-icon">A</div>
                                <div class="color-bar" id="textColorBar"></div>
                            </label>
                        </div>
                        <div class="color-picker-wrapper">
                            <input type="color" id="highlightColor" value="#FFFF00" class="color-input">
                            <label for="highlightColor" class="color-label" title="Highlight Color">
                                <div class="color-icon highlight">🖍</div>
                                <div class="color-bar" id="highlightColorBar"></div>
                            </label>
                        </div>
                    </div>

                    <!-- Alignment Group -->
                    <div class="toolbar-group">
                        <button class="tool-btn" id="alignLeftBtn" title="Align Left">
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor">
                                <rect x="3" y="4" width="18" height="2" />
                                <rect x="3" y="9" width="12" height="2" />
                                <rect x="3" y="14" width="18" height="2" />
                                <rect x="3" y="19" width="12" height="2" />
                            </svg>
                        </button>
                        <button class="tool-btn" id="alignCenterBtn" title="Center">
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor">
                                <rect x="3" y="4" width="18" height="2" />
                                <rect x="6" y="9" width="12" height="2" />
                                <rect x="3" y="14" width="18" height="2" />
                                <rect x="6" y="19" width="12" height="2" />
                            </svg>
                        </button>
                        <button class="tool-btn" id="alignRightBtn" title="Align Right">
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor">
                                <rect x="3" y="4" width="18" height="2" />
                                <rect x="9" y="9" width="12" height="2" />
                                <rect x="3" y="14" width="18" height="2" />
                                <rect x="9" y="19" width="12" height="2" />
                            </svg>
                        </button>
                        <button class="tool-btn" id="alignJustifyBtn" title="Justify">
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor">
                                <rect x="3" y="4" width="18" height="2" />
                                <rect x="3" y="9" width="18" height="2" />
                                <rect x="3" y="14" width="18" height="2" />
                                <rect x="3" y="19" width="18" height="2" />
                            </svg>
                        </button>
                    </div>

                    <!-- List Group -->
                    <div class="toolbar-group">
                        <button class="tool-btn" id="bulletBtn" title="Bullets">
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor">
                                <circle cx="6" cy="6" r="2" />
                                <circle cx="6" cy="12" r="2" />
                                <circle cx="6" cy="18" r="2" />
                                <rect x="10" y="5" width="11" height="2" />
                                <rect x="10" y="11" width="11" height="2" />
                                <rect x="10" y="17" width="11" height="2" />
                            </svg>
                        </button>
                        <button class="tool-btn" id="numberBtn" title="Numbering">
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor">
                                <text x="3" y="8" font-size="8" font-weight="bold">1.</text>
                                <text x="3" y="14" font-size="8" font-weight="bold">2.</text>
                                <text x="3" y="20" font-size="8" font-weight="bold">3.</text>
                                <rect x="10" y="5" width="11" height="2" />
                                <rect x="10" y="11" width="11" height="2" />
                                <rect x="10" y="17" width="11" height="2" />
                            </svg>
                        </button>
                    </div>

                    <!-- Zoom Controls -->
                    <div class="toolbar-group">
                        <button class="tool-btn" id="zoomOutBtn" title="Zoom Out">−</button>
                        <span class="zoom-display" id="zoomLevel">64%</span>
                        <button class="tool-btn" id="zoomInBtn" title="Zoom In">+</button>
                    </div>
                </div>

                <!-- Insert Tab Toolbar (Hidden by default) -->
                <div class="toolbar insert-toolbar">

                    <!-- Text Box -->
                    <div class="toolbar-group">
                        <button class="tool-btn large" id="insertTextBoxBtn">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <rect x="3" y="3" width="18" height="18" rx="2" />
                                <line x1="7" y1="8" x2="17" y2="8" />
                                <line x1="7" y1="12" x2="17" y2="12" />
                                <line x1="7" y1="16" x2="13" y2="16" />
                            </svg>
                            <span>Text Box</span>
                        </button>
                    </div>

                    <!-- Image -->
                    <div class="toolbar-group">
                        <button class="tool-btn large" id="insertImageBtn">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <rect x="3" y="3" width="18" height="18" rx="2" />
                                <circle cx="8.5" cy="8.5" r="1.5" />
                                <path d="M21 15l-5-5L5 21" />
                            </svg>
                            <span>Image</span>
                        </button>
                    </div>

                    <!-- Chart -->
                    <div class="toolbar-group">
                        <button class="tool-btn large disabled" id="insertChartBtn">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <rect x="3" y="12" width="4" height="9" />
                                <rect x="10" y="8" width="4" height="13" />
                                <rect x="17" y="4" width="4" height="17" />
                            </svg>
                            <span>Chart</span>
                        </button>
                    </div>

                    <!-- Table -->
                    <div class="toolbar-group">
                        <button class="tool-btn large disabled" id="insertTableBtn">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <rect x="3" y="3" width="18" height="18" rx="2" />
                                <line x1="3" y1="9" x2="21" y2="9" />
                                <line x1="3" y1="15" x2="21" y2="15" />
                                <line x1="12" y1="3" x2="12" y2="21" />
                            </svg>
                            <span>Table</span>
                        </button>
                    </div>

                    <!-- Video -->
                    <div class="toolbar-group">
                        <button class="tool-btn large disabled" id="insertVideoBtn">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <rect x="2" y="5" width="20" height="14" rx="2" />
                                <polygon points="10 8 16 12 10 16 10 8" fill="currentColor" />
                            </svg>
                            <span>Video</span>
                        </button>
                    </div>

                    <!-- Audio -->
                    <div class="toolbar-group">
                        <button class="tool-btn large disabled" id="insertAudioBtn">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <path d="M9 18V5l12-2v13" />
                                <circle cx="6" cy="18" r="3" />
                                <circle cx="18" cy="16" r="3" />
                            </svg>
                            <span>Audio</span>
                        </button>
                    </div>

                    <!-- Link -->
                    <div class="toolbar-group" style="display: none;">
                        <button class="tool-btn large" id="insertLinkBtn">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" />
                                <path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" />
                            </svg>
                            <span>Link</span>
                        </button>
                    </div>

                    <!-- Comment -->
                    <div class="toolbar-group">
                        <button class="tool-btn large disabled" id="insertCommentBtn">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z" />
                            </svg>
                            <span>Comment</span>
                        </button>
                    </div>
                </div>

                <!-- Shape Format Tab Toolbar (Hidden by default) -->
                <div class="toolbar draw-toolbar">

                    <!-- Shape Fill -->
                    <div class="toolbar-group">
                        <div class="color-picker-wrapper">
                            <input type="color" id="shapeFillColor" value="#4A90E2" class="color-input"
                                title="Fill Color">
                            <label for="shapeFillColor" class="color-label" title="Fill Color">
                                <div class="color-icon">■</div>
                                <div class="color-bar" id="shapeFillColorBar" style="background-color: #4A90E2;"></div>
                            </label>
                        </div>
                    </div>

                    <!-- Shape Outline -->
                    <div class="toolbar-group">
                        <div class="color-picker-wrapper">
                            <input type="color" id="shapeOutlineColor" value="#2E5C8A" class="color-input"
                                title="Outline Color">
                            <label for="shapeOutlineColor" class="color-label" title="Outline Color">
                                <div class="color-icon">□</div>
                                <div class="color-bar" id="shapeOutlineColorBar" style="background-color: #2E5C8A;">
                                </div>
                            </label>
                        </div>
                        <select class="tool-select" id="outlineWidth" title="Outline">
                            <option value="0">No Outline</option>
                            <option value="1">1 pt</option>
                            <option value="2" selected>2 pt</option>
                            <option value="3">3 pt</option>
                            <option value="4">4 pt</option>
                            <option value="5">5 pt</option>
                            <option value="6">6 pt</option>
                        </select>
                    </div>

                    <!-- Effects -->
                    <div class="toolbar-group">
                        <select class="tool-select" id="shapeEffect" title="shapeEffect" style="display: none;">
                            <option value="none">No Effect</option>
                            <option value="shadow">Shadow</option>
                            <option value="reflection">Reflection</option>
                            <option value="glow">Glow</option>
                            <option value="soft-edges">Soft Edges</option>
                            <option value="3d-format">3-D Format</option>
                            <option value="3d-rotation">3-D Rotation</option>
                        </select>
                    </div>

                    <!-- Insert Shapes -->
                    <div class="toolbar-group" style="display: none;">
                        <button class="tool-btn" id="insertRectangle" title="Rectangle">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <rect x="3" y="6" width="18" height="12" rx="2" />
                            </svg>
                        </button>
                        <button class="tool-btn" id="insertCircle" title="Circle">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <circle cx="12" cy="12" r="9" />
                            </svg>
                        </button>
                        <button class="tool-btn" id="insertTriangle" title="Triangle">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <path d="M12 3 L21 20 L3 20 Z" />
                            </svg>
                        </button>
                        <button class="tool-btn" id="insertPentagon" title="Pentagon">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <path d="M12 2 L22 9 L18 21 L6 21 L2 9 Z" />
                            </svg>
                        </button>
                        <button class="tool-btn" id="insertHexagon" title="Hexagon">
                            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor"
                                stroke-width="2">
                                <path d="M12 2 L21 7 L21 17 L12 22 L3 17 L3 7 Z" />
                            </svg>
                        </button>
                    </div>

                    <!-- Arrange -->
                    <div class="toolbar-group">
                        <button class="tool-btn" id="bringToFront" title="Bring to Front">↑↑</button>
                        <button class="tool-btn" id="sendToBack" title="Send to Back">↓↓</button>
                    </div>

                    <!-- Rotate & Flip -->
                    <div class="toolbar-group">
                        <button class="tool-btn" id="rotateLeft90" title="Rotate Left 90°">↶</button>
                        <button class="tool-btn" id="rotateRight90" title="Rotate Right 90°">↷</button>
                        <button class="tool-btn" id="flipVertical" title="Flip Vertical">⇅</button>
                        <button class="tool-btn" id="flipHorizontal" title="Flip Horizontal">⇄</button>
                    </div>

                    <!-- Group/Ungroup -->
                    <div class="toolbar-group" style="display: none;">
                        <button class="tool-btn large" id="groupShapes">
                            <span>Group</span>
                        </button>
                        <button class="tool-btn large" id="ungroupShapes">
                            <span>Ungroup</span>
                        </button>
                    </div>

                    <!-- Align & Distribute -->
                    <div class="toolbar-group">
                        <select class="tool-select" id="alignShapes" >
                            <option value="">Align</option>
                            <option value="left">Align Left</option>
                            <option value="center">Align Center</option>
                            <option value="right">Align Right</option>
                            <option value="top">Align Top</option>
                            <option value="middle">Align Middle</option>
                            <option value="bottom">Align Bottom</option>
                        </select>
                    </div>

                    <!-- Size -->
                    <div class="toolbar-group">
                        <label class="toolbar-label">Size</label>
                        <input type="number" class="tool-input" id="shapeWidth" value="200" min="10" max="1000"
                            placeholder="Width">
                        <span style="color: #b0b0b0;">×</span>
                        <input type="number" class="tool-input" id="shapeHeight" value="100" min="10" max="1000"
                            placeholder="Height">
                    </div>
                </div>

                <!-- Image Format Tab Toolbar (Hidden by default) -->
                <div class="toolbar imageformat-toolbar">
                    <div class="toolbar-group">
                        <button class="tool-btn large">
                            <span>🖼️ Image Tools</span>
                        </button>
                    </div>
                </div>
            </div>

            <!-- Main Content Area -->
            <div class="main-content">

                <!-- Left Sidebar - Slides -->
                <div class="left-sidebar">
                    <div class="sidebar-header">
                        <div class="sidebar-tab active">Slides</div>
                        <div class="sidebar-tab">Outline</div>
                    </div>
                    <div id="slidesContainer" class="slides-container">
                        <!-- Slide previews will be dynamically inserted here -->
                    </div>

                </div>

                <!-- Center - Canvas -->
                <div class="canvas-wrapper">
                    <div class="canvas" id="canvas">

                    </div>
                </div>
            </div>

            <!-- Status Bar -->
            <div class="status-bar">
                <div class="status-left">
                    <span>Slide <span id="currentSlide">1</span> of <span id="totalSlides">3</span></span>
                    <!-- ADD THESE -->
                    <span id="statusMessage" style="margin-left:12px; color:#666;"></span>
                    <span id="loader" style="display:none; margin-left:10px;">⏳</span>
                </div>
                <div class="status-right">
                    <span>English - United States</span>
                </div>
            </div>
        </div>
    </div>

                <!-- inline-editor -->
                <div id="ppt-popupModal" class="modal">
                  <div class="ppt-modal-content">
                    <!-- <img class="inlineeditorcross" id="pptcloseModal" src="<?= base_url('images/Vector2.svg') ?>"> -->
                    <div class="pptmain-content row">
                      <div id="inlineeditoroverlay" class="inlineloader-overlay">
                        <div class="inineloader"></div>
                      </div>
                       
                    <div id="pptcontent">
                  </div>

                </div>
              </div> 

              <div class="row">
                <div class="undo-redo-buttons">
                  <button id="undo" title="Undo"><i class="fa fa-undo" aria-hidden="true"></i></button>
                  <button id="redo" title="Redo"><i class="fa fa-redo" aria-hidden="true"></i></button>
                </div>
              </div> 
            </div>
          </div>
           <!-- inline-editor -->
          <div id="slidecontentdiv" style="overflow-y: <?= !empty($ppt_details) ? 'hidden' : 'auto'; ?>;" >
            <div class="right-side">
              <div class="input-section">
                <input type="text" id="" placeholder="Enter Presentation Name" value='<?php if(isset($_GET['deckid'])) { echo $deck_details->projectname;} ?>'/>
                <button class="download-btn"> 
                <img class="" style="filter:unset;width:13px" id="downloadfaicon" src="<?php echo home_url(); ?>deckbuilder/deckster/images/download-icon-vqw.svg">
                Download</button>
              </div>
              <hr class="divider" />
            </div>

            <div class="theviewstarting"  style="position:relative;height:100%;">
              <?php 
                $slidesorterviewul = '';
                $slidesorterviewli = '';
                $removeactionfromcard = '';
                $blinkdivhide ='flex';
                $slidesorterviewblankdiv = '';
                if (!empty($ppt_details)) {
                  $slidesorterviewul = 'slidesorterviewul';
                  $slidesorterviewli = 'slidesorterviewli';
                  $removeactionfromcard = 'removeactionfromcard';
                  $blinkdivhide = 'none';
                  $slidesorterviewblankdiv = 'slidesorterviewblankdiv';
                }
              ?>
              <?= $this->include('modalInsAll') ?>
              <div id="slideviewsingleslide" style="display: <?= !empty($ppt_details) ? 'flex' : 'none'; ?>;">
                <div class="img-wrapper">
                  
                  <img id="dynamicinlineImage" class="mainviewimg" src="<?= !empty($ppt_details) ? $ppt_details[0]->prod_image : ''; ?>" >
                  <div id="spinnergifcustom" class="inlineloading-spinner" style="display: none;"></div>
                  <div id="spinnerOverlay" class="inlineslide-loading" style="display: none;"></div>
                </div>
              </div>

              <ul id='allSlides' class="slide_scroll image-container <?php echo $slidesorterviewul; ?>"  >
                <?php
                  if (isset($_GET['deckid'])) {
                    $templateCounter=0;
                    $compnaytableforlist = getCompanytable();
                    foreach ($ppt_details as $ppt) { ?>
                      <li slide_table="<?php echo $compnaytableforlist; ?>" slide_id='<?php echo $ppt->id; ?>' class="<?php echo $slidesorterviewli; ?> delete_slideicon image-box li_item<?php echo $templateCounter; ?> onedriveloader_<?php echo $templateCounter; ?> " draggable="true" id="<?php echo $templateCounter;?>" style="position: relative; background-image: url('<?php echo $ppt->prod_image;?>');transition:background-image 1s ease-in-out;">
                        <div class="progress-bar-wrapper">
                          <div class="progress-bar"></div>
                        </div>
                        <div class="delete-icon <?php echo $removeactionfromcard; ?>">
                          <?php
                            $isInlineEditor = '';
                            if ($ppt->inline_enabled == 0) {
                              $isInlineEditor = ' hide ';
                            }
                          ?>        
                          
                          <?php if (true) { ?>
                          <div class=" clubeditoption ppteditoption">
                            <img src="<?= base_url('images/mixicon.svg') ?>" class="delicon pencil" >
                            <div class="editoption peditoption addpadding">
                                <?php if($ppt->inline_enabled == 1){ ?>
                                  <p class="inline_attr_ele_picker removepad" onclick="openFramePPT('<?php echo getPptPath($ppt->id); ?>', 'li_item<?php echo $templateCounter; ?>', '<?php echo $ppt->prod_image; ?>',<?php echo $ppt->id; ?>, this)">
                                    <img src="<?= home_url() ?>deckbuilder/deckster/images/inlineedit.svg">Inline Editor
                                  </p>
                                <?php } ?>
                            </div>
                          </div>
                          <?php } ?>

                          <?php if ($companyId != MOHAWKIND_COMPANY_ID) { ?>
                          <div class="clubeditoption">
                              <img src="<?= base_url('images/inlineedit.svg') ?>" class="delicon pencil" >
                            <div class="editoption">
                              <p onclick="changetocontentui(<?php echo $ppt->id; ?>, '<?php echo $ppt->prod_image; ?>','<?php echo getPptPath($ppt->id); ?>',this)">
                                <img src="<?= base_url('images/inlineedit.svg') ?>">Edit With Neo AI
                              </p>
                            </div> 
                          </div>
                          <?php } ?>

                          <?php if ($companyId != MOHAWKIND_COMPANY_ID) { ?>
                          <div class="clubeditoption">
                              <img src="<?= base_url('images/changedesign.svg') ?>" class="delicon pencil" >
                            <div class="editoption">
                              <p onclick="showDesignOptions(<?= $ppt->id ?>)">
                                <img src="<?= base_url('images/changedesign.svg') ?>">Change Design</p>
                            </div> 
                          </div>
                          <?php } ?>

                          <div class="clubeditoption">
                            <img src="<?= base_url('images/ion_duplicate-outline.svg') ?>" class="delicon pencil" >
                            <div class="editoption">
                              <?php if(true){ ?>
                                <p onclick="copy_template(<?php echo $ppt->id; ?>, '<?php echo getCompanytable(true); ?>', <?php echo $templateCounter;?>,this)">
                                  <img src="<?= home_url() ?>deckbuilder/deckster/images/ion_duplicate-outline.svg">Duplicate Slide
                                </p>
                              <?php } ?>
                            </div> 
                          </div>

                          <div class="clubeditoption" > 
                            <img src="<?= base_url('images/delete.svg') ?>" class="delicon delimg" onclick="deleteSingleSlide(<?php echo $ppt->id; ?>, this)">
                          </div>
                              
                        </div>
                      </li>
                    <?php $templateCounter++; }
                  } 
                ?>
                <div id="blankslide" class="<?php echo $slidesorterviewblankdiv; ?> delete_slideicon image-box-blank blankslide <?php if($companyId != 4){ ?> greyoutplussign <?php } ?>"  onclick="toggleSlidePanel()" >
                  <i class="fa fa-plus addopenlib" aria-hidden="true" ></i>
                </div>
              </ul>
              
              <script>
                // document.getElementById('spinnerOverlay').style.display = 'flex';
                setTimeout(() => {
                  if (typeof createNavigationArrows === 'function') {
                    createNavigationArrows();
                    // document.getElementById('spinnerOverlay').style.display = 'none';
                  }
                }, 5000);

                <?php if (isset($_GET['selectedIds'])) { ?>
                  var currentUrlslidechat = window.location.href;
                  var newUrlslidechat = currentUrlslidechat.replace(/&selectedIds=[^&]*/, '');
                  window.history.pushState({ path: newUrlslidechat }, '', newUrlslidechat);
                  setTimeout(() => {
                    $('#save_download').trigger('click');
                    // document.getElementById('spinnerOverlay').style.display = 'none';
                  }, 5000);
                <?php  } ?>
              </script>
            </div>
          </div>

          <div id="deckProcessingUnit" class="hide deck-container">
            <div class="app-container_1">
              <!-- Main Content -->
              <div class="main-content_1">
                <div class="canvas-toolbar_1">
                  <div class="zoom-controls_1">
                    <button class="zoom-btn_1" onclick="SlideCarousel.zoomOut()">-</button>
                    <div class="zoom-level_1" id="zoomLevel_1">100%</div>
                    <button class="zoom-btn_1" onclick="SlideCarousel.zoomIn()">+</button>
                  </div>
                  <div class="deckapprovecontroll">
                    <div class="slide-indicator_1">
                      <span id="currentSlideNum_1">0</span> / <span id="totalSlides_1">0</span>
                    </div>
                    <a onclick="approveDeck()" id="approveDeckBtn" style="display:none;" >Approve Deck</a>
                  </div>
                </div>
                
                <div class="canvas-area_1">
                  <div class="slide-preview_1" id="slidePreview_1">
                    <div class="loading-container_1" id="loadingContainer_1">
                      <div class="loading-spinner_1"></div>
                      <span>Loading slides...</span>
                    </div>
                    <img src="" alt="Current Slide" class="current-slide_1" id="currentSlide_1" style="display: none;">
                  </div>
                </div>
              </div>

              <!-- Horizontal Carousel -->
              <div class="carousel-container_1">
                <button class="nav-btn_1" onclick="SlideCarousel.previousSlide()" id="prevBtn_1"><i class="fa fa-arrow-left" aria-hidden="true"></i></i></button>
                
                <div class="carousel-track_1">
                  <div class="carousel-slides_1" id="carouselSlides_1">
                    <!-- Slides will be populated by JavaScript -->
                  </div>
                  <div class="carousel-progress_1" id="carouselProgress_1"></div>
                </div>
                
                <button class="nav-btn_1" onclick="SlideCarousel.nextSlide()" id="nextBtn_1"><i class="fa fa-arrow-right" aria-hidden="true"></i></button>
              </div>
            </div>
          </div>
        </div>

        <div class="modal fade" id="modal_template" tabindex="-1" role="dialog" aria-labelledby="generateIconLabel" aria-hidden="true" data-backdrop="static" >
          <div class="modal-dialog modal-lib" role="document">
            <div class="modal-content modal_content_custom">
              <div class="modal-header temp_modal_header">
                <button type="button" class="close" id="templatemodalclose" data-dismiss="modal">
                  <img src="<?= base_url('images/crossicon.svg') ?>" style="width: 14px;height: 14px;position: absolute;right: 1.5rem;top: 1.5rem;" >
                </button>

                <ul class="tabs group" id="tabs">
                  <li class="active"  id="slideTab" onclick="openTab('categoriesfilterdiv', this)">
                    <a>Slide Library</a>
                  </li> 
                  <li id="deckTab" onclick="openTab('allcreateddeck', this)">
                    <a> Deck Library</a>
                  </li> 
                </ul>

                <form class="temp_form" id='serachtempform'>
                  <div class="input-group searchTags">
                    <input type="text" class="form-control form-control-template-search"
                      name="template_name"
                      id="templatesearchid"  placeholder="Search"
                      value="<?php echo htmlspecialchars($searchkey); ?>"
                      onkeypress="searchonEnter(event)"
                      onkeyup="handleBackspace(event)"
                      oninput="toggleClearButton()"
                      onfocus="toggleClearButton()"
                    >
                    <button class="btn btn-primary" id="clearsearchfilter" type="button" onclick="resetallFilter()" >
                      <img src="<?= base_url('images/Vector2.svg') ?>">
                    </button>

                    <button class="btn btn-primary temp_btn" type="button" onclick="searchTemplate()">
                        <i class="fa fa-search"></i>
                    </button>
                  </div>
                </form>
                <a class="resettemplatefilter" href="#" onclick="resetallFilter()">Clear</a>
              </div>

              <div class="panel with-nav-tabs panel-info" style="overflow: hidden;height:92%">
                <div class="panel-body">
                  <div class="tab-content temp_popup">
                    <div class="alert alert-danger text-center d-none" id="errormessageimage1">  Please Select image First </div>
                      <div class="tab-pane fade in active show " id="popular-template">
                        <div class="templatecontentdiv">
                          <div class="filterdiv">

                            <div class="filter-heading">
                              <div class="filtertopcontent">
                                <h2 style="margin-bottom:0px;font-size: 1rem;display: flex;align-items: center;gap: 5px;">
                                <img style="width: 16px;" src="<?= base_url('images/filtericonnew.svg') ?>">Filter</h2>
                                <input type="hidden" class="btn btn-primary" value="Apply Filter" style="background-color:#6667AA;" onclick="apply_templatefilter()">
                              </div>
                            </div>

                            <div id="categories">
                              <div class="subcategory domainsubcat tabcontent" id="categoriesfilterdiv">
                                <input type="text" id="subcategorySearch" onkeyup="filterSubcategories()" placeholder="Search for subcategories...">
                                <p class="catetext" style="margin-bottom:1rem"><img style="width: 18px;" src="<?= base_url('images/categoryppt.svg') ?>"> Select any Category</p>
                                <?php 
                                    foreach ($domaincatfilter as $prod_tag){  ?>
                                      <label class="custom-checkbox subcategory-item">
                                        <span class="checkmarktype"><i class="fa fa-check" aria-hidden="true" ></i></span>   
                                        <input type="checkbox" name="subcategory[]" class="subcategory-checkbox" id="" value="<?php echo $prod_tag->id; ?>" onclick="CategoryChanged(this,true,event,<?php echo $prod_tag->id; ?>)">
                                        <p class="subcat_class domaingroupsubcat" style="text-transform:none"><?php echo $prod_tag->name; ?></p>
                                      </label>
                                <?php } ?>
                              </div>
                              <div class="decksection tabcontent" id="allcreateddeck">  
                                <div class="subcategory domainsubcat deckcategory">
                                  <input type="text" id="mydeckSearch" onkeyup="debouncedFilterDeckSearch()" placeholder="Search by name, creator etc">
                                  <p class="catetext"> <img style="width: 18px;" src="<?= base_url('images/categoryppt.svg') ?>">Select any deck</p>
                                  <div id="deckfiltercat"></div>
                                  <button id="catloadmorebtn" class="deckloadmore">Load More</button>
                                </div>
                              </div>
                            </div>
                          </div>

                          <div class="templatelistingdiv" id="autoscrollproduct">
                            <p id="brandapprovetext">Brand approved slides only</p>
                            <div class="owl-carousel subcat-carousel" id="subcatid"></div>
                            
                            <div id="insertalldiv">
                              <p id="userselecteddeckname"></p>
                              <button style="" type="button" class="insertallbtn" id="insertall" onclick="insertallslide()"> 
                                Insert All
                              </button>
                            </div>

                            <div id="apapendMyTemplates" class='apapendMyTemplates row'></div>
                            <div class="userDecksContainer row"></div>

                            <div class="col-sm-12 text-center">
                              <button id="load-more-btn-edit-ppt" class="btn btn-outline-dark loadmoretemp_btn">Load More</button>
                              <button id="searchloadmorebtn" class="btn btn-outline-dark loadmoretemp_btn">Load More</button>
                              <button id="filterloadmorebtn" class="btn btn-outline-dark loadmoretemp_btn">Load More</button>
                              <button id="pptslideloadmorebtn" class="btn btn-outline-dark loadmoretemp_btn">Load More</button>
                              <button id="deckloadmorebtn" class="btn btn-outline-dark loadmoretemp_btn">Load More</button>
                              <input type="hidden" name="loadmorecount" value="" id="loadmorecount">
                            </div>
                          </div>
                        </div>
                      </div>      
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>

          <div id="popup" class="popup-container">
            <div class="popup-header">
              <h5>Confirm</h5>
            </div>
            <div class="popup-body">
              <p>Project name already exists. If you continue, the previous project will be overwritten.</p>
            </div>
            <div class="popup-footer">
              <a href="#" id="closebtn" onclick="hidePopup()">Cancel</a>
              <a href="#" id="cont" onclick="saveAndDownload()">Continue</a>
            </div>
          </div>
        </div>

      </div>
    </div>
    <!-- </ul> -->

    <div id="overlay1unqppt" class="overlay1" style="display: none;">
      <div class="modal1" >
        <h2 class="modal-text1" style='color:#5c5e9d;text-align:left;font-weight:bold;font-size:19px;margin-bottom: 16px;'>Unsaved changes to presentation</h2>
        <h6 style='text-align: left;color: #aaa;font-weight: 400;'>You have unsaved changes to this presentation. If you leave now, all recent changes will be lost.</h6>
        <div class="button-group1">
          <button onclick="closeThisModalPPT()" class="action-btn1" id="discardppt" style='background-color:#fff;color:#000;padding:6px 20px;'>Cancel</button>
          <button onclick="redirectPPT(true)" class="action-btn1" id='btn_save_project_nameppt' style='background-color:#5c5e9d;padding:6px 20px'>Confirm</button>
        </div>
      </div>
    </div>

    <div id="confirmSync" class="" style="display: none;">
      <div class="modal1" >
      <h2 class="modal-text1" style='color:#5c5e9d;text-align:left;font-weight:bold;font-size:19px;margin-bottom: 16px;font-size:18px'>Sync OneDrive Changes</h2>
      <h6 style='text-align: left;color: #aaa;font-weight: 400;font-size:15px;padding-top:5px;'>We'll sync your OneDrive changes and create a backup of your current work should you want to review or restore any previous changes.</h6>
        <div class="button-group1">
          <button onclick="$('#confirmSync').modal('hide')" class="action-btn1" id="discardpptsync" style='background-color:#fff;color:#000;padding:6px 20px;'>Cancel</button>
          <button onclick="backupDeck(true)" class="action-btn1" id='btn_save_project_nameppt' style='background-color:#5c5e9d;padding:6px 20px'>Confirm</button>
        </div>
      </div>
    </div>

    <div id="restoreModal" class="" style="display: none;">
      <div class="modal1" style="width:70%;">
        <a onclick="$('#restoreModal').modal('hide')" id="restoreclosepopup" style="display: flex;justify-content: end;padding-bottom: 10px;cursor:pointer;">
          <img src="<?= base_url('images/Vector2.svg') ?>" style="filter: grayscale(100%) contrast(0%);width:14px;height:14px">
        </a>
        <h2 class="backuptitle">Backup History</h2>
        <!-- Bootstrap Table inside the Modal Body -->
        <div class="modal-body">
          <table class="backuptable">
            <thead class="backuptablehead">
              <tr>
                <th>Presentation Name</th>
                <th class="center">Slides</th>
                <th class="center">Last Modified</th>
                <th class="center">Action</th>
              </tr>
            </thead>
            <tbody id='restorebody'>
            </tbody>
          </table>
        </div>
        <!-- Buttons -->
        <div class="button-group1"> </div>
      </div>
    </div>
 
    <div id="syncChnagesModal" class="overlay1" style="display: none;">
      <div class="modal1" style='width:51%;padding-top:29px;'>
        <h2 class="modal-text1" style='text-align:left;font-weight:bold;font-size:19px;color:#6467a9;margin-bottom: 16px;'>Sync Changes</h2>
        <h6 style='text-align: left;color: #000;font-weight: bold;font-size: 1.1rem;'>Sync your OneDrive changes.</h6>
        <p style="display: flex;color: #706e6e;"><b>Note: </b> <span style="margin-left:5px"> If not visible, wait a few seconds and retry.</span></p>
        <div class="button-group1">
          <button onclick="$('#syncChnagesModal').hide();localStorage.setItem('isSync',false)" class="action-btn1" id="discardppt" style='background-color:#fff;padding:6px 20px;border: 2px solid #6467a9;color:#000'>Skip Changes</button>
          <button onclick="sync($('#idsyncbtn'))" class="action-btn1" id='syncbutton' style='background-color:#6467a9;padding:6px 20px;color:#fff'>Sync Now</button>
        </div>
      </div>
    </div>

    <!-- Syncing overlay -->
    <div id="syncingoverlay" class="slidepopup-container" style="display: none;">
      <div class="popup-overlay" onclick="closeModalSlide()" style ="text-align: center;padding-top:18%;color:white;">
        <img src="<?= base_url('images/loadinganimiwhite.gif') ?>" alt="Loading..." style="width: 50px; height: 50px;top: 50%; transform: translateY(-50%);">
        <br/>
        Syncing changes from OneDrive...
      </div>
    </div>


    <!-- slide image preview -->
    <div id="image-popup" class="slidepopup-container imageslidepopup">
      <div class="popup-overlay" onclick="closeModalSlide()"></div>
      <div class="slidepopup-content">
        <a onclick="closeModalSlide()" class="closeslidepreview">
          <img src="<?= base_url('images/Vector2.svg') ?>" style="width:12px;height:12px">
        </a>
        <img id="slidepopup-image" src="" alt="Preview Image">
      </div>
    </div>

    <!-- category-subcategory -->
    <div id="assignTemplate" class="modal" style=" ">

      <div class="modal-header">
        <h1 class="mngslidetxt">Manage Slide</h1>
        <a id="cancelBtn"><img style="filter: brightness(0.5);width: 12px;" src="<?= base_url('images/Vector2.svg') ?>" ></a>
      </div>

      <div class="modal-body categorymodal">
        <div class="form-group">
          <label class="form-label">Title</label>
          <input type="text" class="form-control" placeholder="Enter slide title" value="" id="prodTitle">
        </div>

        <label class="form-label">Categories</label>  
        <div id="slideasigncat">

        </div>      
        <div class="form-group catsubcatform">
          <div class="slide-form">
            <div class="catpopupdropdown" style="cursor:pointer" class="dropdown-label" onclick="togglecategorypopup()">
              <p>Select a category</p><i class="fa-solid fa-chevron-down"></i></div>
            <div style="margin-top: 10px;display:none" class="categories-list" id="categoriesDropdown" style="display: none;">
              <?php 
              if (!empty($categories)) {
                foreach($categories as $cat) { ?>
                  <div class="category-item">
                    <input type="checkbox" id="cat-<?= $cat['id']?>" class="category-checkbox selectedCats" value="<?= $cat['id']?>" />
                    <label for="cat-<?= $cat['id']?>"><?= $cat['name']?></label>
                  </div>
                <?php }
              } ?>
            </div>
          </div>
        </div>

        <div id="subcategorySection" style="display: none;">
          <label class="form-label">Subcategories</label> 
          <div id="slideasignsubcat">

          </div>     
          <div class="form-group catsubcatform">
            <div class="slide-form subcate">
              <div class="catpopupdropdown" style="cursor:pointer" class="dropdown-label" onclick="togglesubcategorypopup()">
              <p>Select a Subcategory</p><i class="fa-solid fa-chevron-down"></i></div>
              <div style="margin-top: 10px;display:none" class="categories-list" id="subcategoriesDropdown">
                <?php 
                if (!empty($subcategories)) {
                  $temp = '';
                  foreach($subcategories as $cat) { ?>
                    <div class="hide category-item ctr ctr<?=$cat['category_id']?>">
                      <input type="checkbox" id="subcat-<?= $cat['subcategory_id']?>" class="category-checkbox selectedSubcats" value="<?= $cat['subcategory_id']?>" />
                      <label for="subcat-<?= $cat['subcategory_id']?>"><?= $cat['subcategory_name']?></label>
                    </div>
                    <?php  
                    if ($temp != $cat['category_id']) {
                      $temp = $cat['category_id'];
                    } 
                  }
                } ?>
              </div>
            </div>
          </div>
        </div>
      </div>


      <div class="modal-footer">
        <p class="alert alert-success alert-dismissible fade show" id="changesuccess">
          <strong>Success!</strong> Changes saved successfully.        
        </p>
        <button class="btn" style="background-color:#373434;color:#fff" onclick="saveCatSubcat()">Save Changes</button>
      </div>
    </div>


    <!-- Modal -->
    <div class="modal fade" id="customMembershipModal" tabindex="-1" aria-labelledby="customMembershipLabel" aria-hidden="true">
      <div class="modal-dialog custom-membership-modal-dialog">
        <div class="modal-content custom-membership-modal-content">
          <div class="modal-header">
            <h2 class="modal-text1 commonpopuphtag"  id="customMembershipLabel">Upgrade Your Membership</h2>
            <button type="button" class="btn-close" aria-label="Close" style="background: none !important; border: none !important;" onclick="handleModalClose()">
              <img src="<?= base_url('images/Vector2.svg') ?>" style="width:16px; height:16px">
            </button>
          </div>
          <div class="modal-body">
            <p class="commonpopuparahtag">Get access to premium features by upgrading your membership. Click below to proceed.</p>
          </div>
          <div class="modal-footer">
            <button type="button" class="clspopup" onclick="handleModalClose()">Cancel</button>
            <a href="<?=MEMBERSHIP_URL?>?isEditor=<?=isset($_GET['deckid']) ? $_GET['deckid'] : 0 ?>" class="buyoption clsconvertToPptx" style="background: #EC6F00;border: none;color:#fff;">Proceed to Buy</a>
          </div>
        </div>
      </div>
    </div>

    <!-- theme modal -->
    <div class="modal fade" id="slideModal" tabindex="-1" aria-labelledby="slideModalLabel" aria-hidden="true" data-backdrop="static">
      <div class="modal-dialog modal-dialog-theme">
        <div class="modal-content modal-content-theme ">

          <div class="thememodal-header">
            <h1 class="modal-title fs-5" id="slideModalLabel">Choose Theme</h1>
            <?php
              $grouped_palettes = [];
              foreach ($paletteColors as $color) {
                $id = $color['palette_id'];
                $name = $color['palette_name'];
                if (!isset($grouped_palettes[$id])) {
                  $grouped_palettes[$id] = [
                    'name' => $name,
                    'colors' => []
                  ];
                }
                $grouped_palettes[$id]['colors'][] = $color['hex_code'];
              }
            ?> 
            <div class="palette-dropdown hide">
              <div class="dropdown-button">Colors Palettes ▾</div>
              <div class="palette-list">
                <?php foreach ($grouped_palettes as $palette_id => $palette): ?>
                  <div class="palette-row" data-palette-id="<?= $palette_id ?>">
                    <div class="palette-name"><?= htmlspecialchars($palette['name']) ?></div>
                    <div class="swatch-strip">
                      <?php foreach (array_slice($palette['colors'], 5, 6) as $hex): ?>
                        <div class="swatch" style="background-color: <?= htmlspecialchars($hex) ?>;"></div>
                      <?php endforeach; ?>
                    </div>
                  </div>
                <?php endforeach; ?>
              </div>
            </div>
            <a id="cancelBtnTheme" onclick="closeThemeModal()">
              <img style="filter: brightness(0.5);width: 12px;" src="<?=base_url()?>/images/Vector2.svg">
            </a>
          </div>

          <div class="thememodal-body">
            <div class="row" id="panel1">
              <?php  foreach($themes as $theme) { ?>
                <div class="col-4 mb-4">
                  <div class="slide-item-theme slide-theme-item">
                    <div class="slide-image-theme slide-image-theme slide-1" style="background-image:url('<?=$theme->prod_image?>');">
                      <button class="btn btn-insert-theme" onclick="applydeckTheme('<?=$theme->prod_file?>','<?=$theme->prod_title?>',<?=$theme->id?>)">Apply Theme</button>
                    </div>
                  </div>
                  <div class="theme-title">
                    <p><?= htmlspecialchars($theme->prod_title) ?></p>
                  </div>
                </div>
              <?php } ?>
            </div>

            <div class="theme-panel-2 hide" id="panel2">
            </div>

            <div id="themeoverlay-spinner" style="display:none;">
              <div class="themespinner-background">
                <div class="loader-container" id="loaderContainer">
                  <div class="loader-theme"></div>
                  <div class="timer-theme" id="timer-theme">60s</div>
                </div>
              </div>
            </div>
          </div>

        </div>
      </div>
    </div>

    <div id="beforeOpenOneDrive" >
      <div class="commonpopupbox" >
        <h2 class="modal-text1 commonpopuphtag">Important: Single Tab Usage for File Editing</h2>
        <p class="commonpopuparahtag">You’re opening PowerPoint Online in a new tab. Close the file in other tabs to avoid errors. <br/><br/>If your browser is blocking pop-ups, allow pop-ups for this site to open the slide in Microsoft PowerPoint.</p>

        <div class="button-group1">
          <button onclick="$('#beforeOpenOneDrive').modal('hide')" class="clspopup" id="discardProceedToOneDrive" >Cancel</button>
          <button onclick="$('#beforeOpenOneDrive').modal('hide');downloadDeck(false, 'onedrive', true);" class="clsconvertToPptx" id='proceedToOneDrive'>Confirm</button>
        </div>
      </div>
    </div>


    <div id="onedrive-spinner" style="display:none;">
      <div class="onedrive-background">
        <div class="onedrivepinner"></div>
      </div>
    </div>


    <div id="overlayinlineedit">
      <div id="popupinlineedit">
        <p>You have unsaved changes. Are you sure you want to leave without saving ?</p>
        <div class="buttoninlineeditorpop">
          <button onclick="closeinlinedtiorpopup()" class="clspopup">Close</button>
          <button onclick="saveandcloseinline()" class="clsconvertToPptx">Save</button>
        </div>
      </div>
    </div>
    
    <!-- Your existing popup -->
    <div id="overlayinlineedit1" >
      <div id="popupinlineedit1">
        <button type="button" class="btn-close closesavediscardpopup" aria-label="Close" onclick="closesavediscardpopup()">
          <img src="<?= base_url('images/Vector2.svg') ?>" style="width:12px; height:12px">
        </button>
        <p>You have unsaved changes. Would like to save your presentation or discard them</p>

        <div class="popuppptnaming" >
          <input type="text" id="popupEditTextHeader" placeholder="Untitled Presentation"  >
          <div id="errornamemessage" class="errornamemessage"></div>
        </div>

        <div class="buttoninlineeditorpop1">
          <button onclick="discardChanges()" class="clspopup discard">Discard Changes</button>
          <button onclick="saveandcloseinline1()" class="clsconvertToPptx">Save & Exit</button>
        </div>
      </div>
    </div>

    <div id="overlay1unq" >
      <div class="commonpopupbox" style=''>
        <a onclick="closeThisModal()" style="display: flex;justify-content: end;padding-bottom: 10px;cursor: pointer;">
          <img src="<?= base_url('images/Vector2.svg') ?>" width="16" height="16" >
        </a>
        <h2 class="modal-text1 commonpopuphtag">Create New Deck</h2>
        <p class="commonpopuparahtag">This will open a new window and allow you to create a new deck</p>
        <div class="button-group1">
          <button onclick="closeThisModal()" class="clspopup" >Cancel</button>
          <button onclick="redirectThisModal(true)" class="clsconvertToPptx" id='btn_save_project_name' >Confirm</button>
        </div>
      </div>
    </div>

    <!-- Design Options Modal -->
    <div class="npv-gallery-modal" id="designOptionsModal">
      <div class="npv-gallery-content">
        <div class="npv-gallery-header">
          <p id="designOptionsSubtitle">Select a design variation for Slide</p>
          <button class="npv-gallery-close-btn designpopupclose" onclick="closedesignModal()">×</button>
        </div>
        <div class="npv-grid" id="designOptionsGrid"></div>
        <button class="npv-insert-btn" id="applyDesignBtn" onclick="applySelectedDesign()">Apply Selected Design</button>
      </div>
    </div>

    <div class="npv-gallery-modal" id="npv-galleryModal">
      <div class="npv-gallery-content">
        <div class="npv-gallery-header">
          <button class="npv-gallery-close-btn designpopupclose" onclick="npvCloseGallery()">×</button>
        </div>
        <div class="npv-grid" id="npv-gallery"></div>
        <button class="npv-insert-btn" id="npv-insertBtn" onclick="npvInsertSelected()">Insert Selected</button>
      </div>
    </div>

    <div class="npv-modal" id="npv-modal">
      <div class="npv-modal-content">
        <button class="npv-close-btn" onclick="npvCloseModal()">×</button>
        <img class="npv-modal-image" id="npv-modalImage" src="" alt="Full view">
      </div>
    </div>

  </body>
</html>


<script> 

  window.addEventListener('DOMContentLoaded', () => {
    const saved = localStorage.getItem(restorechathistory); 
    if (saved) {
      const parsed = JSON.parse(saved);
      if (parsed.chatHistorylatest && parsed.chatHistorylatest.length > 0) {
        showtheunifiedchat(); 
        restoreSession();
      }
    }
  });


  $(document).on('click', function(event) {
    if (!$(event.target).closest('.search-container.searchInput').length) {
      $('#searchResults').hide();
    }
  });

  const messageInput = document.getElementById('messageInput');
  const sendButton = document.getElementById('calltounfiedapi');
if (messageInput) {
  messageInput.addEventListener('input', () => {
    const hasText = messageInput.value.trim().length > 0;
    sendButton.disabled = !hasText;
  });
}

  document.addEventListener('keydown', function (event) {
    const activeElement = document.activeElement;
    const isEditable =
      activeElement &&
      (activeElement.tagName === 'INPUT' ||
      activeElement.tagName === 'TEXTAREA' ||
      activeElement.isContentEditable);

    if (isEditable) {
      // Let the browser handle arrow keys normally (move cursor)
      return;
    }

    const slideView = document.getElementById('slideviewsingleslide');
    const isVisible = window.getComputedStyle(slideView).display === 'flex';

    const deckpreviewopen = document.getElementById('previewModal');
    const deckpreviewoVisible = window.getComputedStyle(deckpreviewopen).display === 'none';
    if (!isVisible) return;
    
    if(deckpreviewoVisible){
      if (event.key === 'ArrowLeft') {
        event.preventDefault();
        navigateToPreviousSlide();
      } else if (event.key === 'ArrowRight') {
        event.preventDefault();
        navigateToNextSlide();
      }
    }
  });

  window.addEventListener('DOMContentLoaded', function () {
    clickFirstSlideItem();
  });

  <?php
    $tCounter=0;
    if (!empty($ppt_details)) {
      foreach($ppt_details as $ppt){ ?>
        userSelectedDecks.push(<?php echo basename($ppt->id);?>);
        window.LAST_INSERTED_SLIDE_IDS.push(<?php echo basename($ppt->id);?>);
        imgPaths[<?php echo $tCounter;?>] = {
          prodFile : '<?php echo getPptPath($ppt->id); ?>',
          imgPath : '<?php echo $ppt->prod_image;?>'
        };
        templateCounter++;
        <?php
        $tCounter++;
      } 
    }
  ?>

  function handleStartOver() {
    showthedefaultselectagentview();
    <?php if (isset($_GET['deckid'])): ?>
      if (document.querySelector('.messagebot_1') || document.querySelector('.deck-selection-message_1')) {
        createnewdeckModal('aiagent');
      }
    <?php endif; ?>

    const startoveridbutton = document.getElementById('startoverid');
    if (startoveridbutton) startoveridbutton.style.display = 'none';

    const el = document.querySelector('.deck-appraisal-section_1');
    if (el) {
      el.remove();
    }

    $(".main-banner").show();
    $("#slidecontentdiv").show();
    $("#deckProcessingUnit").addClass("hide");

    const slideType = document.getElementById('slidetypeid');
    const deckSlideType = document.getElementById('deckslidetypeid');
    const slideSelectionModal = document.getElementById('slideSelectionModal');
    
    if (slideType.style.display === 'block') {
      slideType.style.display = 'none';
      deckSlideType.style.display = 'flex';
    }
    
    if (typeof newSession === 'function') {
      newSession();
    }

    if (typeof clearChat === 'function') {
      clearChat();
    }

    if (typeof clearaicontentChat === 'function') {
      clearaicontentChat(); 
    }
  }


  
  function showthedefaultselectagentview(){
    const initialmessagediv = document.getElementById('initialmessagediv');
    const deckSlideType = document.getElementById('deckslidetypeid');
    const unfieddiv = document.getElementById('unfieddiv');
    initialmessagediv.style.display = 'none';

    if (initialmessagediv.style.display === 'none') {
      deckSlideType.style.display = 'none';
      unfieddiv.style.display = 'flex';
    }
  }

  function openneopopup() {
    document.querySelector('.neo-popup').style.display = 'block';
  }

  function closeneopopup() {
    document.querySelector('.neo-popup').style.display = 'none';
  }

  function startneo(){
    var agendaViewagent = document.getElementById('agendaview');
    closeneopopup();
    if (!agendaViewagent.offsetParent) {
      agendacreation();
    }
  }

  $('#agentgetSearchTemplate').on('input', function() {
    if ($(this).val() === '') {
      $('#agentclearSearch1').css({'display': 'none'}); 
      $('#searchslideanddeckicon').css({'display': 'block'}); 
    } else {
      $('#agentclearSearch1').css({'display': 'block'}); 
      $('#agentsearchslideanddeckicon').css({'display': 'none'}); 
    }
  });

  var agentSearchTemplate = document.getElementById('agentgetSearchTemplate');
  if (agentSearchTemplate) {
    agentSearchTemplate.addEventListener('keypress', function(event) {
      if (event.key === 'Enter') {
        agendacreation();
        var agentValue = this.value;
        document.getElementById('getSearchTemplate').value = agentValue;
        $('#searchslideanddeckicon').css({'display' : 'none'});
        $('#clearSearch1').css({'display' : 'block'});
        $('#content2 .product-container').html('');
        $('#content1 .product-container').html('');
        decksOffset = 0;
        slidesOffset = 0;
        getSearchTemplate();
      }
    });
  }

  var dropdownButtoncolorpalette = document.querySelector('.dropdown-button');
  if (dropdownButtoncolorpalette) {
    dropdownButtoncolorpalette.addEventListener('click', function() {
      var dropdown = document.querySelector('.palette-dropdown');
      if (dropdown) {
        dropdown.classList.toggle('show-dropdown');
      }
    });
  }

  // Select swatch
  document.querySelectorAll('.swatch').forEach(swatch => {
    swatch.addEventListener('click', function() {
      document.querySelectorAll('.swatch').forEach(s => s.classList.remove('selected'));
      this.classList.add('selected');
      document.querySelector('.dropdown-button').click();
    });
  });

</script>

<!-- Load jQuery first if not already loaded in header -->
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/owl.carousel/dist/owl.carousel.min.js"></script>


<!-- Load your main application scripts with defer for better performance -->
<script defer src="<?= base_url('assets/js/editor.js') ?>"></script>
<script defer src="<?= base_url('assets/js/unifieddeckslide.js') ?>"></script>
<script defer src="<?= base_url('assets/js/aicontent.js') ?>"></script>
<script defer src="<?= base_url('assets/js/deckagent.js') ?>"></script>
<script defer src="<?= base_url('assets/js/inlineeditorai.js') ?>"></script>

<script defer src="<?= base_url('assets/js/editor-core.js') ?>"></script>
<script defer src="<?= base_url('assets/js/home-tab.js') ?>"></script>
<script defer src="<?= base_url('assets/js/insert-tab.js') ?>"></script>
<script defer src="<?= base_url('assets/js/shape-format.js') ?>"></script>
<script defer src="<?= base_url('assets/js/image-tab.js') ?>"></script>

<script>
  let pendingLinkHref = null;
  
  function openInlineEditPopup() {
    document.getElementById("overlayinlineedit1").style.display = "flex";
  }

  function closemembershipaction(){
    const iconElement = document.getElementById("downloadIcon");
    const textElement = document.getElementById("downloadText");
    if (iconElement && textElement) {
      iconElement.style.width = "16px";
      iconElement.style.height = "16px";
      iconElement.src = base_url + "images/downloaddecksnew.svg";
      textElement.textContent = "Download";
    }
  }

  function handleModalClose() {
    closemembershipaction();
    $('#customMembershipModal').modal('hide');
  }

  function closeinlinedtiorpopup1() {
    document.getElementById("overlayinlineedit1").style.display = "none";
    pendingLinkHref = null;
  }

  function saveandcloseinline1() {
    const inputField = document.getElementById('popupEditTextHeader');
    const errorMessage = document.getElementById('errornamemessage');
    const inputVal = inputField.value.trim();

    if (!inputVal) {
      errorMessage.textContent = 'Please provide a presentation name';
      inputField.classList.add('input-error');
      inputField.focus();
      return; // Stop save
    }

    errorMessage.textContent = '';
    inputField.classList.remove('input-error');

    document.getElementById('edittextheader').value = inputVal;
    $('#save_download').trigger('click');
    document.getElementById('overlayinlineedit1').style.display = 'none';

    if(newdeckcall){
      newdeckcall=false;
      createnewdeckModal();
      return;
    }
    
    if(isDownloadPending){
      downloadDeck();
    }

    if(isopeninonedrivepending){
      downloadDeck(false, 'onedrive');
    }

    if(gethelpexitsprojectname){
      saveDownload(false,'sendService');
    }

    if (pendingLinkHref) {
      window.location.href = pendingLinkHref;
    }
  }

  function closesavediscardpopup() {
    closemembershipaction();
    document.getElementById('overlayinlineedit1').style.display = 'none';
  }

  document.getElementById('popupEditTextHeader').addEventListener('input', function () {
    document.getElementById('errornamemessage').textContent = '';
    this.classList.remove('input-error');
  });

  // This handles the Discard action
  async function discardChanges() {
    await removethedeckanditsslide();
    document.getElementById("overlayinlineedit1").style.display = "none";
    if ( pendingLinkHref && pendingLinkHref.trim() !== "" && pendingLinkHref.trim().toLowerCase() !== "javascript:void(0)" 
          && pendingLinkHref.trim() !== "#" && !pendingLinkHref.trim().toLowerCase().startsWith("javascript:")) {
      window.location.href = pendingLinkHref;
    }
    else{
      var workspaceLinkHref = "<?= base_url('workspace') ?>";
      window.location.href = workspaceLinkHref;
    }
  }

  function removethedeckanditsslide() {
     const deckid = "<?php echo isset($_GET['deckid']) ? $_GET['deckid'] : ''; ?>";
    return new Promise((resolve, reject) => {
      $.ajax({
        url: `ajax/removedeckfromdatabase`,
        method: 'POST',
        data: { deckidtoremoved: deckid },
        success: function (response) {
          // console.log(response);
          resolve(response); // resolves the promise
        },
        error: function (xhr, status, error) {
          reject(error); // rejects the promise
        }
      });
    });
  }

  // Intercept all anchor (<a>) clicks
  document.addEventListener('click', function (e) {
    const anchor = e.target.closest('a');
    if (anchor && anchor.href && anchor.id !== 'newdeckbtnid' && anchor.id !== 'btnSyncOneDrive' && anchor.id !== 'applythemetag'  && anchor.id !== 'toggleViewBtn' && anchor.id !=='syncdb' && anchor.id !=='applycolortag') {
      e.preventDefault();
      if (isInsertingSlides) {
        return; 
      }

      pendingLinkHref = anchor.href;
      const headerInput = document.getElementById('edittextheader');
      const headerValue = headerInput ? headerInput.value.trim() : '';

      if (headerValue === 'Untitled Presentation') {
        const slideList = document.getElementById('allSlides');
        const slideItems = slideList.querySelectorAll('li');

        if(slideItems){
          let noofslidecount = slideItems.length;
        }

        if(noofslidecount>0){
          openInlineEditPopup();
        }
        else{
          discardChanges();
        }
      } else {
        window.location.href = pendingLinkHref;
      }
    }
  });

  function assignThemetoDeck(selectedthemeid){
    const deckid = "<?php echo isset($_GET['deckid']) ? $_GET['deckid'] : ''; ?>";
    const themeId = selectedthemeid;
    return new Promise((resolve, reject) => {
      $.ajax({
        url: `ajax/assignThemetoDeck`,
        method: 'POST',
        data: { 
          deckidtoupdate : deckid,
          appliedtheme  : themeId        
        },
        success: function (response) {
          resolve(response); // resolves the promise
        },
        error: function (xhr, status, error) {
          reject(error); // rejects the promise
        }
      });
    });
  }

  function checkThemeonDeck(templateId){
    const deckid = currentdeckid;
    return new Promise((resolve, reject) => {
      $.ajax({
        url: `ajax/getThemeAssigntoDeck`,
        method: 'POST',
        data: { 
          deckid : deckid,
          templateId : templateId    // <-- SEND TEMPLATE ID
        },
        success: resolve,
        error: reject
      });
    });
  }

  async function applyThemeSlide(data) {
    // createLoaderPlaceholder();
    return new Promise((resolve, reject) => {
      $.ajax({
        url: `ajax/processThemeSlides`,
        type: 'POST',
        data: {
          'fileName' :data.template_file, 
          'themeUrl' :data.theme_file,
        },
        success: async function(response) {
          // console.log("themeappliedsuccess",response.newSlideIds);
          if (!response.newSlideIds || response.newSlideIds.length === 0) {
            console.warn("No new slides returned from API");
            resolve(response);
            return;
          }
          // const templateId = response.newSlideIds;
          
          for (const slideId of response.newSlideIds) {
            await apply_template(tablename, slideId);
          }

          $('#save_download').trigger('click');
          resolve(response);
        },
        error: function(xhr, status, error) {
          console.error("Error merging slides:", error);
          reject(error); // Reject the promise with the error
        }
      });
    });
  }

  document.addEventListener('DOMContentLoaded', function() {
    console.log('DOMContentLoaded 8 fired');
    const seeAllButton = document.getElementById('seeAllButton');
    const agentModalOverlay = document.getElementById('agentModalOverlay');
    const agentCloseButton = document.getElementById('agentCloseButton');

    // Open agent modal
    if(seeAllButton){
      seeAllButton.addEventListener('click', function() {
        document.querySelectorAll('.slide-item').forEach(item => {
          const text = item.textContent.trim();
          if (!['Agenda Slide', 'Roadmap Slide', 'Timeline Slide','Title Slide','Thank You Slide','Problem Solution Slide','Feature Benefits Slide','Testimonial Slide','Pros & Cons Slide','Break Slide','Chart Slide'].includes(text)) {
            item.classList.add('disabled');
          }
        });
        agentModalOverlay.classList.add('active');
        document.body.style.overflow = 'hidden';
      });
    }

    // Close agent modal
    function closeAgentModal() {
      agentModalOverlay.classList.remove('active');
      document.body.style.overflow = 'auto';
    }

    agentCloseButton.addEventListener('click', closeAgentModal);

    // Close agent modal when clicking overlay
    agentModalOverlay.addEventListener('click', function(e) {
      if (e.target === agentModalOverlay) {
        closeAgentModal();
      }
    });

    // Close agent modal with Escape key
    document.addEventListener('keydown', function(e) {
      if (e.key === 'Escape' && agentModalOverlay.classList.contains('active')) {
        closeAgentModal();
      }
    });

    // Handle slide item clicks
    const slideItems = document.querySelectorAll('.slide-item');
    slideItems.forEach(item => {
      item.addEventListener('click', function() {
        closeAgentModal();
      });
    });

    // Handle show more/less clicks
    const showToggleButtons = document.querySelectorAll('.show-toggle');
    showToggleButtons.forEach(button => {
      button.addEventListener('click', function() {
        const categoryElement = this.closest('.agentcategory');
        const hiddenItems = categoryElement.querySelectorAll('.slide-item.hidden');
        
        if (this.textContent === 'Show more...') {
          // Show hidden items
          hiddenItems.forEach(item => {
            item.classList.remove('hidden');
          });
          this.textContent = 'Show less...';
        } else {
          // Hide items beyond first 5
          const allItems = categoryElement.querySelectorAll('.slide-item');
          allItems.forEach((item, index) => {
            if (index >= 5) {
              item.classList.add('hidden');
            }
          });
          this.textContent = 'Show more...';
        }
      });
    });
  });

  function clearSearchInput(){
    console.log("poopo");
  }
</script>

<script type="text/javascript">
  function waitForElements() {
    const btn = document.getElementById('addShapeBtn');
    const dropdown = document.getElementById('shapeDropdown');
    
    if (btn && dropdown) {
      btn.onclick = function(e) {
        e.preventDefault();
        e.stopPropagation();
        if (dropdown.style.display === 'block') {
          dropdown.style.display = 'none';
        } else {
          dropdown.style.display = 'block';
        }
      };
      
      // Handle shape clicks
      const options = dropdown.querySelectorAll('.shape-option');
      options.forEach(option => {
        option.onclick = function(e) {
          e.preventDefault();
          const shapeType = this.getAttribute('data-shape');
          dropdown.style.display = 'none';
        };
      });
      
      // Close on outside click
      document.onclick = function(e) {
        if (!e.target.closest('.add-shape-container')) {
          dropdown.style.display = 'none';
        }
      };

    } else {      
      // Try again in 1 second
      setTimeout(waitForElements, 1000);
    }
  }

  // Start checking for elements
  waitForElements();
</script>
