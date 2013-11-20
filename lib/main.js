var parser = require("xml2json");
var admzip = require("adm-zip");
var fs = require("fs");

var docx = module.exports = function() {};
var seperateRules = {
    // body : {
    //     open : "<w:body>",
    //     close : "</w:body>",
    //     counter : 0
    // },
    paragraph : {
        open : "<w:p [^>]+>",
        close : "</w:p>",
        counter : 0,
        type : "paragraph"
    }
};

var removeRules = {
    wlang : {
        open : "<w:lang [^/]+/>",
        type : "lang"
    },
    pStyle : {
        open : "<w:pStyle [^/]+/>",
        type : "style"
    },
    rFonts : {
        open : "<w:rFonts [^/]+/>",
        type : "fonts"
    },
    t : {
        open : "(<w:t>|<w:t [^>]+>)",
        close : "</w:t>",
        type : "t"
    },
    rPr : {
        open : "(<w:rPr>|<w:rPr [^>]+>)",
        close : "</w:rPr>",
        type : "rpr"
    },
    pPr : {
        open : "(<w:pPr>|<w:pPr [^>]+>)",
        close : "</w:pPr>",
        type : "ppr"
    },
    drawing : {
        open : "(<w:drawing>|<w:drawing [^>]+>)",
        close : "</w:drawing>",
        type : "drawing"
    },
    pinline : {
        open : "(<wp:inline>|<wp:inline [^>]+>)",
        close : "</wp:inline>",
        type : "inline"
    },
    pcNvGraphicFramePr : {
        open : "(<wp:cNvGraphicFramePr>|<wp:cNvGraphicFramePr [^>]+>)",
        close : "</wp:cNvGraphicFramePr>",
        type : "cNvGraphicFramePr"
    },
    agraphicData : {
        open : "(<a:graphicData>|<a:graphicData [^>]+>)",
        close :"</a:graphicData>",
        type : "graphData"
    },
    agraphic : {
        open : "(<a:graphic>|<a:graphic [^>]+>)",
        close : "</a:graphic>",
        type : "graphic"
    },
    picpic : {
        open : "(<pic:pic>|<pic:pic [^>]+>)",
        close : "</pic:pic>",
        type : "pic"
    },
    picnvPicPr : {
        open : "(<pic:nvPicPr>|<pic:nvPicPr [^>]+>)",
        close : "</pic:nvPicPr>",
        type : "nvPicPr"
    },
    astretch : {
        open : "(<a:stretch>|<a:stretch [^>]+>)",
        close : "</a:stretch>",
        type : "stretch"
    },
    picblipfill : {
        open : "(<pic:blipFill>|<pic:blipFill [^>]+>)",
        close : "</pic:blipFill>",
        type : "blipFill"
    },
    ablipr : {
        open : "<a:blipr(.*?)[^>]+>",
        close : "</a:blipr>",
        type : "blipr"
    },
    cNvPicPr : {
        open : "(<pic:cNvPicPr>|<pic:cNvPicPr [^>]+>)",
        close : "</pic:cNvPicPr>",
        type : "cNvPicPr"
    },
    aext : {
        open : "(<a:ext>|<a:ext (.*?)[^>]+>)",
        close : "</a:ext>",
        type : "aext"
    },
    aextLst : {
        open : "(<a:extLst>|<a:extLst (.*?)[^>]+>)",
        close : "</a:extLst>",
        type :"aextlst"
    },
    ablip : {
        open : "(<a:blip>|<a:blip (.*?)[^>]+>)",
        close : "</a:blip>",
        type : "ablip"
    },
    spPr : {
        open : "(<pic:spPr>|<pic:spPr (.*?)[^>]+>)",
        close : "</pic:spPr>",
        type : "spPr"
    },
    xfrm : {
        open : "(<a:xfrm>|<a:xfrm (.*?)[^>]+>)",
        close : "</a:xfrm>",
        type : "xfrm"
    },
    off : {
        open : "(<a:off>|<a:off (.*?)[^>]+>)",
        close : "</a:off>",
        type : "offset"
    },
    prstGeom : {
        open : "(<a:prstGeom>|<a:prstGeom (.*?)[^>]+>)",
        close : "</a:prstGeom>",
        type : "offset"
    },
    ln : {
        open : "(<a:ln>|<a:ln (.*?)[^>]+>)",
        close : "</a:ln>",
        type : "ln"
    }
};

var simpleRemove = {
    r : {
        open : "(<w:r>|<w:r [^>]+>)",
        close : "</w:r>",
        type : "r"
    },
    noProof : {
        open : "(<w:noProof/>|<w:noProof [^/]+/>)",
        type : "proof"
    },
    pextend : {
        open : "(<wp:extent/>|<wp:extent [^/]+/>)",
        type : "extend"
    },
    peffectExtent : {
        open : "(<wp:effectExtent/>|<wp:effectExtent [^/]+/>)",
        type : "effectExtent"
    },
    agraphicFrameLocks : {
        open : "(<a:graphicFrameLocks/>|<a:graphicFrameLocks (.*?)[^/]+/>)",
        type : "graphicFrameLocks"
    },
    anofill : {
        open : "(<a:noFill/>|<a:noFill [^/]+/>)",
        type : "nofill"
    },
    apiclocks : {
        open : "(<a:picLocks/>|<a:picLocks (.*?)[^/]+/>)",
        type : "piclocks"
    },
    afillrect : {
        open : "(<a:fillRect/>|<a:fillRect (.*?)[^/]+/>)",
        type : "fillRect"
    },
    aext : {
        open : "(<a:ext/>|<a:ext (.*?)[^/]+/>)",
        type : "aext"
    },
    uselocaldpi : {
        open : "(<a[0-9]+:useLocalDpi/>|<a[0-9]+:useLocalDpi (.*?)[^/]+/>)",
        type : "useLocalDpi"
    },
    srcRect : {
        open : "(<a:srcRect/>|<a:srcRect (.*?)[^/]+/>)",
        type : "srcRect"
    },
    off : {
        open : "(<a:off/>|<a:off (.*?)[^/]+/>)",
        type : "offset"
    },
    avlst : {
        open : "(<a:avLst/>|<a:avLst (.*?)[^/]+/>)",
        type : "avlst"
    },
    cNvPr : {
        open : "(<pic:cNvPr/>|<pic:cNvPr (.*?)[^/]+/>)",
        type : "img"
    },
    bookmarkStart : {
        open : "(<w:bookmarkStart/>|<w:bookmarkStart (.*?)[^/]+/>)",
        type : "bookmarkStart"
    },
    bookmarkEnd : {
        open : "(<w:bookmarkEnd/>|<w:bookmarkEnd (.*?)[^/]+/>)",
        type : "bookmarkEnd"
    }
};

var markers = {
    pic : {
        content : "(?:<wp:docPr (.*?)[^/]+/>)",
        lookup : "(?:<wp:docPr(.*?)id=\")[0-9]+",
        remove : "(?:<wp:docPr(.*?)id=\")",
        type : "pic"
    }
};

docx.prototype.getContent = function(fn) {
    fs.readFile('./.tmp/docx-content/word/document.xml', 'utf8', function(err, data){
        if(err) throw err;
        fn(data);
    });
};

docx.prototype.parseXML = function(data) {
    var json = {};
    for(var i in seperateRules) {
        var patt = new RegExp("(?:"+seperateRules[i].open+")(.*?)(?="+seperateRules[i].close+")","g");
        json[seperateRules[i].type] = [];
        while((match = patt.exec(data)) !== null) {
            if(seperateRules[i].type == "paragraph") {
                json[seperateRules[i].type].push(match[0].replace(new RegExp(seperateRules[i].open,""),""));
            }
        }
    }
    for(var j in json.paragraph) {
        json.paragraph[j] = this.cleanParse(json.paragraph[j]);
    }
    for(var x in json.paragraph) {
        json.paragraph[x] = this.simpleRemove(json.paragraph[x]);
    }
    console.log(json.paragraph);
    return json.paragraph.join("\n");
};

docx.prototype.simpleRemove = function(data) {
    for(var i in simpleRemove) {
        data = data.replace(new RegExp(simpleRemove[i].open,"g"),"");
        if(simpleRemove[i].close)
            data = data.replace(new RegExp(simpleRemove[i].close,"g"),"");
    }
    for(var x in markers) {
        var patt = new RegExp(markers[x].content,"");
        var match = patt.exec(data);
        if(match !== null) {
            if(markers[x].type == "pic") {
                var tmp = data.replace(patt,"");
                var lookup = new RegExp(markers[x].lookup,"");
                var remove = new RegExp(markers[x].remove,"");
                var r_lookup = lookup.exec(match[0]);
                var id = r_lookup[0].replace(remove,"");
                data = {
                    wrap : tmp,
                    id : id,
                    type : "pic"
                };
            }
        }
    }
    return data;
};

docx.prototype.cleanParse = function(data) {
    var self = this;
    for(var i in removeRules) {
        if(removeRules[i].close === undefined) {
            var singlePatt = new RegExp(removeRules[i].open,"g");
            data = data.replace(singlePatt,"");
        } else {
            var patt = new RegExp("(?="+removeRules[i].open+")(.*?)(?="+removeRules[i].close+")","g");
            while((match = patt.exec(data)) !== null) {
                var temp = match[0].replace(new RegExp(removeRules[i].open,""),"");
                data = data.replace(new RegExp("(?="+removeRules[i].open+").*?"+removeRules[i].close,""),temp);
                //console.log(temp)
            }
        }
    }
    return data;
};

docx.prototype.parse = function(file) {
    var zip = new admzip(file);
    zip.extractAllTo("./.tmp/docx-content/",true);
    return this.init();
};

docx.prototype.init = function(){
    var self = this;
    this.getContent(function(data){
        //console.log(parser.toJson(data));
        var result = self.parseXML(data);
        fs.writeFile('./.tmp/result.txt',result, function(err){
            if(err) throw err;
            console.log("saved");
        });
    });
};