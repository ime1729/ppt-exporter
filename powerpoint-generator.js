var og=require('officegen');
var fs=require('fs');
var ppt=officegen('pptx');
var slide=ppt.makeNewSlide();
slide.addText("Title goes here",
              {y:20,x:0,cx:'100%',cy:60,font_size:48,align:'center'});
var out=fs.createWriteStream('out.pptx');
ppt.generate(out);
//console.log('This works!');
function genast(lines){
    var ret=[];
    var curSlide={};
    lines.forEach(function(line){
        function updateBody(head,body){
            if(head===null){
                if(curSlide.content===undefined){
                    throw new Error("Line: \""+body+"\" not a continuation");}
                else{curSlide.content[curSlide.content.length-1]
                             .command+=("\n"+body);}}
            else{
                if(curSlide.content===undefined){
                    curSlide.content=[{parse:head,command:body}];}
                else{curSlide.content.push({parse:head,command:body});}}}
        var pos=line.indexOf(' ');
        var head=line.substr(0,pos);
        var tail=line.substr(pos+1);
        switch(head.toLowerCase()){
        case "#title":curSlide.title=tail;break;
        case "#author":curSlide.author=tail;break;
        case "#chart":console.log("TODO");break;
        case "#slide":ret.push(curSlide);curSlide={title:tail};break;
        case "#text":updateBody(head,tail);break;
        case "#chart":updateBody(head,tail);break;
        case "#":updateBody(null,tail);break;
        case "":break;
        default:
            throw new Error("Unknown command:"+head);}});
    ret.push(curSlide);
    ret.forEach(function(slide){if(slide.content!=undefined){
        var jcon=[];//JSON content parsed from commands
        slide.content.forEach(function(e){
            if(e.parse=="#text"){jcon.push({text:e.command});}
            else if(e.parse=="#chart"){jcon.push(function(p){
                var s=p.split(" ");
                //only works if there is one dataset (which is satisfactory
                // for now)
                return {type:s.shift(),set:s.shift(),
                        cat:s.shift().slice(1,-1).split(","),
                        label:s.join(" ")};}(e.command));}
            else{throw new Error("Unknown element");}});}});
    return ret;}
module.exports.genast=genast;
