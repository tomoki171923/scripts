#!/bin/awk -f

BEGIN {
  split("", csvData); 
  split("", pulseData);
  split("", lstKey);
  lstKey[1] = "pulseId";
  lstKey[2] = "startTime";
  lstKey[3] = "interval";
  lstKey[4] = "renderTime";
  lstKey[5] = "Nodes rendered";
  lstKey[6] = "Layout Pass";
  lstKey[7] = "CSS Pass";
  lstKey[8] = "Update bounds";
  lstKey[9] = "Waiting for previous rendering";
  lstKey[10] = "Copy state to render graph";
  lstKey[11] = "Dirty Opts Computed";
  lstKey[12] = "Render Roots Discovered";
  lstKey[13] = "Painting";
  for(i=1; i<=13; i++) {
	  if(i != 1) {
		  printf(",");
	  }
	  printf(lstKey[i]);
  }
  printf("\n");
  startTime=0;
}

function formatDate(t) {
	return sprintf("%s.%03d", strftime("%H:%M:%S", t / 1000 + 15 * 3600), t % 1000);
}

function printPulseCsv(pulseData) {
	startTime += pulseData["interval"];
	printf("%d,%s",
		   pulseData["pulseId"],
		   formatDate(startTime));
	for(i=3; i<=13; i++) {
		printf(",%d", pulseData[lstKey[i]]);
	}
	printf("\n");
}

{
	if(match($0, /T.*\+([0-9]+)ms\): (.*)$/, grps)) {
		pulseData[grps[2]] += grps[1];
	} else if(match($0, /(Nodes rendered): ([0-9]+)/, grps)) {
		pulseData[grps[1]] = grps[2];
	} else if(match($0, /PULSE: ([0-9]+) \[([\-0-9]+)ms:([\-0-9]+)ms\]/, grps)) {
		if("pulseId" in pulseData) {
			printPulseCsv(pulseData);
			split("", pulseData);
		}
		pulseData["pulseId"] = grps[1];
		pulseData["interval"] = grps[2];
		pulseData["renderTime"] = grps[3];
	} else {
	    while(match($0, /\[([0-9]+) ([\-0-9]+)ms:([\-0-9]+)ms\](.*)/, grps)) {
			if("pulseId" in pulseData) {
				printPulseCsv(pulseData);
				split("", pulseData);
			}
			pulseData["pulseId"] = grps[1];
			pulseData["interval"] = grps[2];
			pulseData["renderTime"] = grps[3];
			$0 = grps[4];
		}
	}
}

END {
	if("pulseId" in pulseData) {
		printPulseCsv(pulseData);
	}
}
