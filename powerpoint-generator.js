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
        case "#chart":updateBody(head,body);break;
        case "#":updateBody(null,tail);break;
        case "":break;
        default:
            throw new Error("Unknown command:"+head);}
        /*console.log("Line:"+head+"||"+tail);
        console.log(ret);
        console.log(curSlide);*/
    });
    ret.push(curSlide);
    console.log(ret);
    return ret;
}
module.exports.genast=genast;
