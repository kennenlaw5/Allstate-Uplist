function counter(range) {
  //Created By Kennen Lawrence 720-317-5427
  //Version 1.1
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getActiveSheet();
  range=sheet.getRange(2,2,sheet.getLastRow()-1,4).getValues();
  var names=[[]];var found=false;var num=0;
  var rank=[];var temp;var r=0;
  for(var i=0;i<range.length;i++){if(range[i][0]!=undefined&&range[i][0]!=""&&range[i][0]!=null){found=true;}}
  if(!found){return "No Data";}
  found=false;
  names[0][0]=range[0][0].toLowerCase();
  if(range[0][3].toLowerCase()=="yes"||range[0][3].toLowerCase()=="yes "||range[0][3].toLowerCase()==" yes"){names[0][1]=1;}else{names[0][1]=0;}
  for(var i=1;i<range.length;i++){
    if(range[i][0]!=""){
      if(range[i][0]!="Client Advisor"){
        found=false;
        for(var j=0;j<names.length && found==false;j++){
          if(range[i][0].toLowerCase()==names[j][0]){
            found=true;
            if(range[i][3].toLowerCase()=="yes"||range[i][3].toLowerCase()=="yes "||range[i][3].toLowerCase()==" yes"){
              names[j][1]+=1;
            }
          }
        }
        if(found==false){
          num=names.length;
          names[num]=[range[i][0].toLowerCase(),0];
          if(range[i][3].toLowerCase()=="yes"||range[i][3].toLowerCase()=="yes "||range[i][3].toLowerCase()==" yes"){names[num][1]+=1;}
          Logger.log(names[num]);
        }
      }
    }else{i=range.length;}
  }
  var totalYes=0;
  var totalNames=names.length;
  for(i=0;i<totalNames;i++){
    totalYes+=parseInt(names[i][1]);
  }
  for(i=0;i<totalNames;i++){
    if(r==0){rank[0]=names[0];}
    else{
      for(var m=0;m<rank.length;m++){
        if(names[i][1]>rank[m][1]){
          temp=rank[m];
          rank[m]=names[i];
          names[i]=temp;
        }
      }
      rank[r]=names[i];
    }
    r+=1;
  }
  for(i=0;i<totalNames;i++){
    names[i]=rank[i][0];
    names[i]=names[i].split(" ");
    names[i][0]=names[i][0][0].toUpperCase() + names[i][0].substring(1);
    names[i][1]=names[i][1][0].toUpperCase()+names[i][1].substring(1);
    rank[i][0]=names[i][0]+" "+names[i][1];
  }
  Logger.log(totalNames);
  Logger.log(totalYes);
  Logger.log(rank);
  //sheet.getRange(1,5,totalNames,2).setValues(rank);
  return rank;
}

function agentRatio(values) {
  /* Code left for debugging!
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getSheetByName('BMW August 2018');
  values=sheet.getRange(2,3,sheet.getLastRow()-1,3).getValues();*/
  var agents=[];var found=false;
  for(var i=0;i<values.length;i++){
    if(values[i][0]!=undefined && values[i][0]!=null && values[i][0]!=""){
      values[i][0]=values[i][0].split(" ")[0];
    }
  }
  for(i=0;i<values.length;i++){
    if(values[i][0]!=undefined && values[i][0]!=null && values[i][0]!="" && values[i][0].toLowerCase()!="lsp"){
      if(agents.length==0){
        agents[0]=[values[i][0].toLowerCase(),0,0];
        if(values[i][2].toLowerCase()=="yes"){agents[0][1]+=1;}
        else if(values[i][2].toLowerCase()=="no"){agents[0][2]+=1;}
      }
      else{
        found=false;
        for(var j=0;j<agents.length;j++){
          if(values[i][0].toLowerCase()==agents[j][0]){
            found=true;
            if(values[i][2].toLowerCase()=="yes"){agents[j][1]+=1;}
            else if(values[i][2].toLowerCase()=="no"){agents[j][2]+=1;}
          }
        }
        if(!found){
          agents[agents.length]=[values[i][0].toLowerCase(),0,0];
          if(values[i][2].toLowerCase()=="yes"){agents[agents.length-1][1]+=1;}
          else if(values[i][2].toLowerCase()=="no"){agents[agents.length-1][2]+=1;}
        }
      }
    }
  }
  for(i=0;i<agents.length;i++){
    agents[i][0]=agents[i][0][0].toUpperCase()+agents[i][0].substring(1);
  }
  Logger.log(agents);
  return agents;
}