/// <reference path="jquery/jquery.js" />
/// <reference path="ww.jquery.min.js" />
/// <reference path="highlightjs/highlight.pack.js" />

// global page reference
window.helpBuilder = null;

(function () {
    // interface
    helpBuilder = {
        initializeLayout: initializeLayout,
        initializeTOC: initializeTOC,
        isLocalUrl: isLocalUrl,        
        expandTopic: expandTopic,        
        expandParent: expandParents,
        tocExpandAll: tocExpandAll,
        tocExpandTop: tocExpandTop,
        tocCollapseAll: tocCollapseAll,
        tocClearSearchBox: tocClearSearchBox,
        highlightCode:  highlightCode,
        updateDocumentOutline: updateDocumentOutline,
        refreshDocument: refreshDocument,
        fox: null,
        configureAceEditor: null, // set in aceConfig,
        searchIndex: null,
        activeId: null
    };  
   

    function initializeLayout(notused) {        
       // for old IE versions work around no FlexBox
        if (navigator.userAgent.indexOf("MSIE 9") > -1 ||
	        navigator.userAgent.indexOf("MSIE 8") > -1 || 
	        navigator.userAgent.indexOf("MSIE 7") > -1)
            $(document.body).addClass("old-ie");

        // modes: none/0 - with sidebar,  1 no sidebar
        var mode = getUrlEncodedKey("mode");
        if (mode)
            mode = mode * 1;
        else
            mode = 0;

        // Legacy processing page=TopicId urls to load topic by id
        var page = getUrlEncodedKey("page");
        if (page)
            loadTopicAjax(page);

        var isLocal = isLocalUrl();
        
        $(".page-content").on("click", "a", function (e) {            
            var href = $(this).attr("href");                

            // ajax navigation online
            if (!isLocal && href.startsWith("_") || href == "index.htm") {      
                loadTopicAjax(href);
                return false; // stop navigation
            } 
            // external links open in new window
            if ( href.startsWith("http://") || href.startsWith("https://") )
            {                                
                if (helpBuilder.fox) {                                        
                    if (window.helpBuilder.fox.navigatepreviewurl(href))
                        return false;   // done - no bubbling                    
                }
                window.open(href,"_blank");
                return false; // done no bubbling
            }

            return true;
        });



        if (!isLocal){
	        // load internal help links via Ajax
	        $(".page-content").on("click", "a", function (e) {            
                var href = $(this).attr("href");                
	            if (href.startsWith("_")) {      
	                loadTopicAjax(href);
	                return false; // stop navigation
	            } 
	        });

            var id = getIdFromUrl();
            if (id){
                setTimeout(function() {
                    $(".toc li a").removeClass("selected");
                    var $a = $("#" + id);
                    $a.addClass("selected");
                    if ($a.length > 0)                    
                        $a[0].scrollIntoView(); 
                },100);
            }
    	}

        if (isLocalUrl() || mode === 1) {
            hideSidebar();                        
        } else {
            $.get("tableofcontents.htm", loadTableOfContents);

            // sidebar or hamburger click handler
            $(document.body).on("click", ".sidebar-toggle", toggleSidebar);
            $(document.body).on("dblclick touchend", ".splitter", toggleSidebar);
             
            $(".sidebar-left").resizable({
                handleSelector: ".splitter",
                resizeHeight: false
            });

            // handle back/forward navigation so URL updates
            window.onpopstate = function (event) {                
                if (history.state && history.state.URL)
                    loadTopicAjax(history.state.URL,true);
            }             
            
        }

        $(".main-content").scroll(debounce(scrollSpy, 100));

        timeToRead();                     

        setTimeout(function() {
            helpBuilder.refreshDocument();
            scrollSpy();            
            
        },10);
    }

    function loadTopicAjax(href, noPushState) {
        var hrefPassed = true;
        
        if(window.innerWidth < 768)
            hideSidebar();

        if (typeof href != "string") {
            var $a = $(this);
            href = $a.attr("href");
            hrefPassed = false;
            
            $(".toc li a").removeClass("selected");
            $a.addClass("selected");               
        }

        var id = href.replace(".htm","");     
        expandTopic(id);

        if (helpBuilder.activeId == id)
            return;
        
        helpBuilder.activeId = id;           
        
        if ($(this).parent().find("i.fa").length > 0) {
            var searchVal = $("#SearchBox").val();            
            if (!searchVal)    {        
                expandParents(href) ;                             
            }
        }

        // ajax navigation
        if (href.startsWith("_") || href == "index.htm") {
            $.get(href, function (html) {                
                clearSearchPane();

                var $html = $(html);

                var title = html.extract("<title>", "</title>");
                window.document.title = title;

                var $content = $html.find(".main-content");
                if ($content.length > 0) {
                    html = $content.html();
                    $(".main-content").html(html);                    

                    // update the navigation history/url in addressbar
                    if (window.history.pushState && href.startsWith('_')) {
                        if (!noPushState)  
                            window.history.pushState({ title: '', URL: href }, "", href);                        

                        $(".selected").removeClass("selected");                            
                        $("#" + id).addClass("selected");
                        $("#SearchBox").val('');                        
                        expandParents(id);
                    }
                    else                     
                        $(".main-content").scrollTop(0);
                } else
                    return;

                var $banner = $html.find(".banner");
                if ($banner.length > 0);
                $(".banner").html($banner.html());

                helpBuilder.refreshDocument();

                $(".main-content").scroll(debounce(scrollSpy,100));
                scrollSpy();
                
   
            });
            return false;  // don't allow click
        }
        return true;  // pass through click
    }; 

    // Initial load of the TOC via XHR
    function loadTableOfContents(html) {
        var $tocContent = $("<div>" + getBodyFromHtmlDocument(html) + "</div>").find(".toc-content");
        $("#toc").html($tocContent.html());

        showSidebar();

        // handle AJAX loading of topics        
        $(".toc").on("click", "li a", loadTopicAjax);

        initializeTOC();

        
        setTimeout(function() { 
            expandTopic("index"); 

            if (!window.lunr)
                $("#SearchDetailed").remove();
        },200);
        return false;
    }

    // initialization of TOC either on first load or reload (from cache)
    function initializeTOC() {

        // if running in frames mode link to target frame and change mode
        if (window.parent.frames["wwhelp_right"]) {
            $(".toc li a").each(function () {
                var $a = $(this);
                $a.attr("target", "wwhelp_right");
                var a = $a[0];
                a.href = a.href + "?mode=1";
            });
            $("ul.toc").css("font-size", "1em");
        }

        // Handle clicks on + and -
        $("#toc").on("click","li>i.fa",function () {            
            expandTopic($(this).find("~a").prop("id") );                        
        });
        $("#toc").on("click","#SearchBoxClearButton",helpBuilder.tocClearSearchBox);

        $("#SearchBox").focus();
    
        var page = getUrlEncodedKey("page");
        if (page) {
            page = page.replace(/.htm/i, "");
            expandParents(page);
        }
        if (!page) {
            page = window.location.href.extract("/_", ".htm");
            if (page)
                expandParents("_" + page);
        }

        var topic = getUrlEncodedKey("topic");
        if (topic) {
            var id = findIdByTopic();
            if (id) {
                var link = document.getElementById(id);
                var id = link.id;
                expandTopic(id);
                expandParents(id);
                loadTopicAjax(id + ".htm");
            }
        }

        
        var searchKeyupFunc = debounce(function() {
            var $searchBox = $(this);
            var searchText = $searchBox.val();

            var $toc = $(".toc.topic-tree");
            var $searchPane = $(".toc.search-results");

            var mode = $("#SearchAdvanced").val();            
            if (searchText) 
               if(mode.toLowerCase() === "simple")
                simpleSearch(searchText);            
               else
               advancedSearch(searchText); 
            else 
                clearSearchPane();
            
        },200);
        $("#SearchBox").on("keyup",searchKeyupFunc);
    }

    var sidebarTappedTwice = false;
    function toggleSidebar(e) {

        // handle double tap
        if (e.type === "touchend" && !sidebarTappedTwice) {
            sidebarTappedTwice = true;
            setTimeout(function () { sidebarTappedTwice = false; }, 300);
            return false;
        }
        var $sidebar = $(".sidebar-left");
        var oldTrans = $sidebar.css("transition");
        $sidebar.css("transition", "width 0.5s ease-in-out");
        if ($sidebar.width() < 20) {
            $sidebar.show();
            $sidebar.width(400);
        } else {
            $sidebar.width(0);
        }

        setTimeout(function () { $sidebar.css("transition", oldTrans) }, 700);
        return true;
    }


    function clearSearchPane() {
        var $toc = $(".toc.topic-tree")
        var $searchPane = $(".toc.search-results");

        $toc.show();
        $searchPane.hide();
        $searchPane.html('');
    }
    function simpleSearch(searchText) {
        var $searchPane = $(".toc.search-results");
        $searchPane.html('');
        $searchPane.show();
        $(".toc.topic-tree").hide();                     
        
        var results = $(".toc.topic-tree li>a:containsNoCase(" + searchText + ")");            
        if (results.length == 0)
            return;
        
        if(results.length > 200) {
            $li = $("<li>Too many matches: " + results.length + " - please narrow your search.</li>");
            $searchPane.append($li);
            return;
        }


        $searchPane.append("<li class='search-result-header' >Search Results <div class='badge badge-info badge-super'>" +  results.length +"</div></li>");         

        setTimeout(function() {                           
            for (let index = 0; index < results.length; index++) {
                const $li = $(results[index]).parent();
                const $li2 = $li.clone();                             
                $li2.find("ul").remove();
                $li2.find(".fa").remove();
                $searchPane.append($li2);                
            }
        },1);
    }
    function advancedSearch(searchText) {
        if (!lunr) return;

        var $searchPane = $(".toc.search-results");
        $searchPane.html('');
        $searchPane.show();
        
        $(".toc.topic-tree").hide();                     

        if (!helpBuilder.searchIndex) {
            var pxhr = $.getJSON("SearchIndex.js")
                .done( function (documents) {

                    // build the index
                    var idx = lunr(function () {
                        this.ref('id')
                        this.field('keywords')
                        this.field('title')
                        this.field('body')
                        this.field('id')

                        this.metadataWhitelist = ['id']

                        documents.forEach(function (doc) {
                            this.add(doc)
                        }, this)
                    })
                    console.dir(idx);

                    // save the index and docs so we don't have to reload for each search
                    helpBuilder.searchIndex = idx;
                    searchIndex(searchText);
                })
                .fail(function() { 
                    $("#SearchOptions").hide();
                    clearSearchPane();
                    return;
                });
        }else 
            searchIndex(searchText);

        function searchIndex(searchText) {
                var res = helpBuilder.searchIndex.search(searchText);    
                $searchPane.append("<li class='search-result-header' >Search Results <div class='badge badge-info badge-super'>" +  res.length +"</div></li>");         
                       
                for (let index = 0; index < res.length; index++) {
                    var id =  res[index].ref;
                    var $a = $("#" + id.toLowerCase() );
                    var $li = $a.parent();         
                    const $li2 = $li.clone();                             
                    $li2.find("ul").remove();
                    $li2.find(".fa").remove();
                    $searchPane.append($li2);                                                                
                }
        }
    }

    function hideSidebar() {
        var $sidebar = $(".sidebar-left");
        var $toggle = $(".sidebar-toggle");
        var $splitter = $(".splitter");
        $sidebar.hide();
        $toggle.hide();
        $splitter.hide();
    }
    function showSidebar() {
        var $sidebar = $(".sidebar-left");
        var $toggle = $(".sidebar-toggle");
        var $splitter = $(".splitter");
        $sidebar.show();
        $toggle.show();
        $splitter.show();
    }
    
    function expandTopic(topicId) {        
        var $href = $("#" + topicId.replace(".htm", ""));

        var $ul = $href.next();
        $ul.toggle();

        var $button = $href.prev().prev();

        if ($ul.is(":visible"))
            $button.removeClass("fa-caret-right").addClass("fa-caret-down");
        else
            $button.removeClass("fa-caret-down").addClass("fa-caret-right");
    }

    function expandParents(id, noFocus) {
        if (!id)
            return;

        var $node = $("#" + id.toLowerCase());
        $node.parents("ul").show();

        if (noFocus)
            return;

        var node = $node[0];
        if (!node)
            return;

        //node.scrollIntoView(true);
        node.focus();
        setTimeout(function () {
            window.scrollX = 0;
        });

    }
    function findIdByTopic(topic) {
        if (!topic) {
            var query = window.location.search;
            var match = query.search("topic=");
            if (match < 0)
                return null;
            topic = query.substr(match + 6);
            topic = decodeURIComponent(topic);
        }
        var id = null;
        $("a").each(function () {
            if ($(this).text().toLowerCase() == topic.toLocaleLowerCase()) {
                id = this.id;
                return;
            }
        });
        return id;
    }
    
    function tocClearSearchBox() {  
        var val = $("#SearchBox").val();       
        if (!val)
            return;  // already empty
    
        $("#SearchBox").val("");
        clearSearchPane();

        // make all visible
        $(".toc li").show();

        tocCollapseAll();
                
        setTimeout(function() {
            // make sure we preserve selection
            var $el = $(".selected");
            var id = '';
            if ($el.length > 0)
                id = $el[0].id;
            
            if (id)
                expandParents(id,false);

            $("#SearchBox")[0].focus();
        },150);
    }

    function tocCollapseAll() {

        $("ul.toc > li ul:visible").each(function () {
            var $el = $(this);
            var $href = $el.prev();            

            var $ul = $href.next();
            $ul.toggle();
    
            var $button = $href.prev().prev();    
            if ($ul.is(":visible"))
                $button.removeClass("fa-caret-right").addClass("fa-caret-down");
            else
                $button.removeClass("fa-caret-down").addClass("fa-caret-right");
        });
    }

    function tocExpandAll() {
        $("ul.toc > li ul:not(:visible)").each(function () {
            var $el = $(this);
            var $href = $el.prev();

            var $ul = $href.next();
            $ul.toggle();
    
            var $button = $href.prev().prev();    
            if ($ul.is(":visible"))
                $button.removeClass("fa-caret-right").addClass("fa-caret-down");
            else
                $button.removeClass("fa-caret-down").addClass("fa-caret-right");    
        });
    }
    function tocExpandTop() {        
        $("ul.toc>li>ul:not(:visible)").each(function () {
            var $el = $(this);
            var $href = $el.prev();
            var id = $href[0].id;
            expandTopic(id);
        });
    }
    function isLocalUrl(href) {
        if (!href)
            href = window.location.href;

        return href.startsWith("mk:@MSITStore") ||
	           href.startsWith("file://")
    }
    function getIdFromUrl(href) {
        if (!href)
            href = window.location.href;

        if(!href.startsWith("_")) {
            href = href.extract("/_", ".htm");
            if(href)
                href = "_" + href;
        }
        
        if (href.startsWith("_"))
            return href.toLowerCase().replace(".htm","");
        
        return null;
    }
    function mtoParts(address, domain, query) {
        var url = "ma" + "ilto" + ":" + address + "@" + domain;
        if (query)
            url = url + "?" + query;
        return url;
    }


    function highlightCode() {   
        var pres = document.querySelectorAll("pre>code");
        for (var i = 0; i < pres.length; i++) {
            hljs.highlightBlock(pres[i]);
        }

        if (window.highlightJsBadge)
            window.highlightJsBadge();
    }

    function CreateHeaderLinks() {
        var $h3 = $(".content-body>h2,.content-body>h3,.content-body>h4,.content-body>h1");

        $h3.each(function () {            
            var $h3item = $(this);
            $h3item.css("cursor", "pointer");

            var tag = $h3item[0].id; //text().replace(/\s+/g, "");

            var $a = $("<a />")
	            .attr({
	                name: tag,
	                href: "#" + tag
	            })
	            .addClass('link-icon')
	            .addClass('link-hidden')
                .attr('title', 'click this link and set the bookmark url in the address bar.');

            $h3item.prepend($a);

            $h3item
	            .hover(
	                function () {
	                    $a.removeClass("link-hidden");
	                },
	                function () {
	                    $a.addClass("link-hidden");
	                })
	            .click(function () {
	                window.location = $a.prop("href");
	            });
        });

        if(location.hash){
            setTimeout(function() {
                // navigate # links if any            

                var hash = location.hash.replace("#","");
                var el$ = $(location.hash+ ",a[name=" + hash + "]");            
                if (el$.length > 0)
                {
                    el$[0].scrollIntoView(true);
                    var mc$ = $(".main-content");                
                    mc$[0].scrollTop = mc$[0].scrollTop - 80;
                }
            });
        }
    }

    function updateDocumentOutline(){
        var navbar$ = $(".topic-outline-content");                       
        navbar$.html("");

        var headers$ = $(".content-pane").find("h1,h2,h3,h4");
        
        if (headers$.length < 2)
        {                   
            $(".content-pane").removeClass("topic-outline-visible");
            $(".topic-outline-header").hide(false);
            return;
        }                

        for (var index = 0; index < headers$.length; index++) {
            var el = headers$[index];
            var id = el.id;
            if (!id) {
                el.id = safeId(el.innerText);
                id = el.id
            }
                
            var space = "";
            if (el.nodeName == "H1")
                space = "outline-level1";
            else if (el.nodeName == "H2")
                space = "outline-level2";
            else if (el.nodeName == "H3")
                space = "outline-level3";
            else if (el.nodeName == "H4")
                space = "outline-level4";
            var a$ = $("<a></a>")
                .prop("href", "#" + id)
                .text(el.innerText);
            if (space)
                a$.addClass(space);

            navbar$.append(a$);
        } 

        $(".content-pane").addClass("topic-outline-visible");
        $(".topic-outline-header").show(true);                    
    }

    function scrollSpy() {        
        var headers$ = $(".topic-outline-content>a");        
        if(headers$.length < 1)
            return;

        for (var index = 0; index < headers$.length; index++) {
            const hd$ = $(headers$[index]);
            var id = hd$.attr('href');

            var id$;
            try{
                 id$ = $(id);
            }catch(ex) {
                continue;
            }
            if(id$.length < 1)
                continue;

            if(id$.isInViewport())
            {                
                $(".topic-outline-content *").removeClass("active");
                hd$.addClass("active");
                break;
            }
        }
    }

    $.fn.isInViewport = function() {
        var elementTop = $(this).offset().top;
        var elementBottom = elementTop + $(this).outerHeight();
        var viewportTop = $(window).scrollTop();
        var viewportBottom = viewportTop + $(window).height();
        return elementBottom > viewportTop && elementTop < viewportBottom;
    };   

    function safeId(inputString) {
        if (!inputString) return inputString;
       var id =  $.trim(inputString)
            .replace(/-/g,"--")
            .replace(/[\s,-,:,.,\',\",\\,/,(,),#<,$,%,@,!,*']/g,"-");
        return id;               
    }

    /* 
        Updates the document with post-processing scripts.
        Called when page reloads.
    */
    function refreshDocument() {
        helpBuilder.highlightCode();
        
        timeToRead();
        CreateHeaderLinks();
        
        helpBuilder.updateDocumentOutline();            
    }
    
    /*
     *  timeToRead()
     * 
     *  Time To Read figures writes out the `Time To Read` text at the top
     *  of the document by injecting it into the `#TimeToRead` element at 
     *  the top of content templates.
     * 
     *  This function delegates to `localizedReadingTimeText()` to localize
     *  the text displayed optionally which can be overridden with a custom
     *  global (at window.) `localizedReadingTimeText()` function to provide
     *  additional localizations.
    */    
    function timeToRead() {
        ttr$ = $("#TimeToRead");
        if (ttr$.length == 0)
            return;
        
        var content = $('.content-pane').text();
        
        var regExWords = /\s+/gi;
        var wordCount = content.replace(regExWords, ' ').split(' ').length;     
        var wordsPerMinute = 250;   //  assumed avg reading speed
        
        // figure out minutes to read
        var minutes = 0;
        if (wordCount >= wordsPerMinute)
           minutes = wordCount / wordsPerMinute;           
        minutes = minutes.toFixed(0);
        
        // Language: en, fr, de, es etc. and then localize if possible to browser langauge
        var lang = navigator.language.substr(0,2);        
        var readingTimeText = localizedReadingTimeText(minutes, lang);
        
        if (readingTimeText)
            ttr$.html('<span><i class="fa fa-clock-o"></i> '+ readingTimeText +'</span>');
    }    

    /*
     * localizedReadingTimeText(minutes, lang)
     *   
     *  Parameters:  
     *   minutes  -   Minutes it takes to read
     *   lang     -   Language as a two letter string (en, de, fr) 
     *                based on browser's active language
     * 
     * Function that 
     * 
     * Can be overriden at the `window` level by creating a custom
     * function called localizedReadingTimeText() that returns a string
     * with the same signature as shown here. If that function returns
     * a string that string is used, otherwise this function continue
     * to run using default functionality.      
    */
    function localizedReadingTimeText(minutes, lang){      
        var readingTimeText = '';   // this value will get set

        // Check to see if a global localizedReadingTimeText function exists
        if (window.localizedReadingTimeText) {
           readingTimeText = window.localizedReadingTimeText(minutes, lang);
           if (readingTimeText)
              return readingTimeText;
        }


        if(lang == "de") {
            if (minutes > 0) {
              
                if (minutes == 1) readingTimeText = 'ca. 1 Minute zu lesen';
                else if (minutes < 9) readingTimeText = 'ca. ' + minutes + ' minuten zum lesen';
                else if (minutes < 13) readingTimeText = 'ca. 10 Minuten zum lesen';
                else if (minutes < 18) readingTimeText = 'ca. 15 Minuten zum lesen';
                else if (minutes < 23) readingTimeText = 'ca. 20 Minuten zum lesen';
                else if (minutes < 28) readingTimeText = 'ca. 25 Minuten  zum lesen';
                else if (minutes < 38) readingTimeText = 'ca. eine halbe Stunde zum lesen';
                else if (minutes < 50) readingTimeText = 'ca. 45 Minuten zum lesen';
                else if (minutes < 70) readingTimeText = 'ca. eine Stunde zum lesen';
                else if (minutes < 80) readingTimeText = 'ca. eine Stunde und 15 Minuten zum lesen';
                else if (minutes < 100) readingTimeText = 'ca. eineinhalb Stunden zum lesen';
                else if (minutes < 130) readingTimeText = 'ca. zwei Stunden zum lesen';                                
                else readingTimeText = 'mehr als zwei Stunden zum lesen';
            }
            else
                readingTimeText = 'weniger als eine Minute zum lesen';            
        }
        else if(lang == "fr") {
            if (minutes > 0) {
              
                if (minutes == 1) readingTimeText = 'environ 1 minute pour lire';
                else if (minutes < 9) readingTimeText = 'environ ' + minutes + ' minutes pour lire';
                else if (minutes < 13) readingTimeText = 'environ 10 minutes pour lire';
                else if (minutes < 18) readingTimeText = 'environ 15 minutes pour lire';
                else if (minutes < 23) readingTimeText = 'environ 20 minutes pour lire';
                else if (minutes < 28) readingTimeText = 'environ 25 minutes pour lire';
                else if (minutes < 38) readingTimeText = 'environ une demi-heure pour lire';
                else if (minutes < 50) readingTimeText = 'environ 45 minutes pour lire';
                else if (minutes < 70) readingTimeText = 'environ 1 heure pour lire';
                else if (minutes < 80) readingTimeText = 'environ 1 heure et 15 minutes pour lire';
                else if (minutes < 100) readingTimeText = 'environ 1 heure et demie pour lire';
                else if (minutes < 130) readingTimeText = 'environ 2 heure pour lire';                                
                else readingTimeText = 'plus de deux heures pour lire';
            }
            else
                readingTimeText = 'moins d\'une minute pour lire';            
        }
        else if(lang == "es") {
            if (minutes > 0) {
              
                if (minutes == 1) readingTimeText = 'environ 1 minute pour lire';
                else if (minutes < 9) readingTimeText = 'environ ' + minutes + ' minutes pour lire';
                else if (minutes < 13) readingTimeText = 'unos 10 minutos para leer';
                else if (minutes < 18) readingTimeText = 'unos 15 minutos para leer';
                else if (minutes < 23) readingTimeText = 'unos 20 minutos para leer';
                else if (minutes < 28) readingTimeText = 'unos 25 minutos para leer';
                else if (minutes < 38) readingTimeText = 'aproximadamente media hora para leer';
                else if (minutes < 50) readingTimeText = 'unos 45 minutos para leer';
                else if (minutes < 70) readingTimeText = 'aproximadamente 1 hora para leer';
                else if (minutes < 80) readingTimeText = 'aproximadamente 1 hora y 15 minutos para leer';
                else if (minutes < 100) readingTimeText = 'aproximadamente una hora y media para leer';
                else if (minutes < 130) readingTimeText = 'aproximadamente 2 horas para leer';                                
                else readingTimeText = 'más de 2 horas para leer';
            }
            else
                readingTimeText = 'menos de 1 minuto para leer';            
        }
        else {
            // English is the default
            if (minutes > 0) {         
                if (minutes == 1) readingTimeText = 'about 1 minute to read';
                else if (minutes < 9) readingTimeText = 'about ' + minutes + ' minutes to read';
                else if (minutes < 13) readingTimeText = 'about 10 minutes to read';
                else if (minutes < 18) readingTimeText = 'about 15 minutes to read';
                else if (minutes < 23) readingTimeText = 'about 20 minutes to read';
                else if (minutes < 28) readingTimeText = 'about 25 minutes to read';
                else if (minutes < 38) readingTimeText = 'about a half hour to read';
                else if (minutes < 50) readingTimeText = 'about 45 minutes to read';
                else if (minutes < 70) readingTimeText = 'about one hour to read';
                else if (minutes < 80) readingTimeText = 'about an hour and 15 minutes to read';
                else if (minutes < 100) readingTimeText = 'about an hour and a half to read';
                else if (minutes < 130) readingTimeText = 'about two hours to read';                                
                else readingTimeText = 'more than two hours to read';
            }
            else
                readingTimeText = 'less than 1 minute to read';            
        }   

        return readingTimeText;
    }

})();



// global functions called from HelpBuilder Dev
function updatedocumentcontent(html, pragmaLine) {
    // don't render invalid HTML that includes body tags
    if (html.indexOf("<body>") > -1)
        return;

    $(".main-content").html(html);

    // refresh syntax coloring and header links
    helpBuilder.refreshDocument();
                
    if (typeof pragmaLine === "number")
        scrolltopragmaline(pragmaLine);
    
}

function scrolltopragmaline(lineno) {             
    if (typeof lineno != "number")   
        return;
    
    $mc = $(".main-content");               
    if (lineno < 2){
            $mc[0].scrollTop = 0;
            return;
    }

    try {
        var $el = $("#pragma-line-" + lineno);
                    
        if ($el.length < 1) {
            var origLine = lineno;
            for (var i = 0; i < 3; i++) {
                lineno++;
                $el = $("#pragma-line-" + lineno);
                if ($el.length > 0)
                    break;
            }
            if ($el.length < 1) {
                lineno = origLine;
                for (var i = 0; i < 3; i++) {
                    lineno--;
                    $el = $("#pragma-line-" + lineno);
                    if ($el.length > 0)
                        break;
                }
            }
            if ($el.length < 1)
                return;
        }
        
        $el.addClass("line-highlight");
        setTimeout(function() { $el.removeClass("line-highlight"); }, 1200);

        setTimeout(function() {
            $el[0].scrollIntoView(); 
            if (lineno > 2)                             
                $mc[0].scrollTop = $mc[0].scrollTop - 80;                                   
        });
        
        
    }
    catch(ex) {  
        
    }       
}

function initializeinterop(fox) 
{
    window.helpBuilder.fox = fox; 
}

/* ES5 POLYFILLS */

// String.trim() for ES5 polyfill
if (!String.prototype.trim) {
    String.prototype.trim = function () {
        return this.replace(/^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g, '');
    };
}

if (!String.prototype.trimEnd) {
    String.prototype.trimEnd = function (c) {
        if (c)
            return this.replace(new RegExp(c.escapeRegExp() + "*$"), '');
        return this.replace(/\s+$/, '');
    };
}

if (!String.prototype.startsWith) {
    Object.defineProperty(String.prototype, 'startsWith', {
        value: function (search, rawPos) {
            pos = rawPos > 0 ? rawPos | 0 : 0;
            return this.substring(pos, pos + search.length) === search;
        }
    });
}

// Object.assign() for ES5 polyfill
if (typeof Object.assign !== 'function') {
    // Must be writable: true, enumerable: false, configurable: true
    Object.defineProperty(Object, "assign", {
        value: function assign(target, varArgs) { // .length of function is 2
            'use strict';
            if (target === null || target === undefined) {
                throw new TypeError('Cannot convert undefined or null to object');
            }

            var to = Object(target);

            for (var index = 1; index < arguments.length; index++) {
                var nextSource = arguments[index];

                if (nextSource !== null && nextSource !== undefined) {
                    for (var nextKey in nextSource) {
                        // Avoid bugs when hasOwnProperty is shadowed
                        if (Object.prototype.hasOwnProperty.call(nextSource, nextKey)) {
                            to[nextKey] = nextSource[nextKey];
                        }
                    }
                }
            }
            return to;
        },
        writable: true,
        configurable: true
    });
}

