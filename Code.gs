function counter() {
  //Created By Kennen Lawrence 720-317-5427
  //Version 1.0
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getActiveSheet();
  var range=sheet.getRange(2,2,sheet.getLastRow()-1,4).getValues();
  var names=[[]];var found=false;var num=0;
  var rank=[];var temp;var r=0;
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
