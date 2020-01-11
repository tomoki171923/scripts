BEGIN {
	FS = ",";
	timeSpan = 0.1; # •ª
	timeEnd = timeSpan * 60;
	timeEndStr = timeToStr(timeEnd);
	lastRenderCnt = 0;
	renderInterval = 0;
	renderIntervalCnt = 0;
	
	printf("endTime,Nodes rendered,ApInterval,RenderInterval,LayoutPass,CSSPass,Updatebounds,Waitingforpreviousrendering,Copystatetorendergraph,DirtyOptsComputed,RenderRootsDiscovered,Painting\n");

	resetCounters();
}

function resetCounters() {
	pulseCnt = 0;
	renderCnt = 0;

	nodes = 0;
	interval = 0;
	layoutPass = 0;
	cssPass = 0;
	updateBounds = 0;
	waiting = 0;
	copyState = 0;
	dirtyOpts = 0;
	renderRoot = 0;
	painting = 0;
	renderInterval = 0;
	renderIntervalCnt = 0;
}

function timeToStr(secs) {
	return sprintf("%s.000", strftime("%H:%M:%S", secs + 15 * 3600));
}

function printAvgs() {
	if (pulseCnt > 0 && renderCnt > 0) {
		printf("%s,%d,%d,%3.1f,%3.1f,%3.1f,%3.1f,%3.1f,%3.1f,%3.1f,%3.1f,%3.1f\n",
		   timeEndStr,
		   nodes / renderCnt,
		   interval / pulseCnt,
		   renderInterval / renderIntervalCnt,
		   layoutPass / renderCnt,
		   cssPass / renderCnt,
		   updateBounds / renderCnt,
		   waiting / renderCnt,
		   copyState / renderCnt,
		   dirtyOpts / renderCnt,
		   renderRoot / renderCnt,
		   painting / renderCnt);
	}
}

$1 ~ /^[0-9]+$/ {
	interval += $3;
	if(lastRenderCnt > 0) {
		renderInterval += $3;
		renderIntervalCnt += 1;
	}
	if($2 >= timeEndStr) {
		printAvgs();
		resetCounters();
		timeEnd += timeSpan * 60;
		timeEndStr = timeToStr(timeEnd);
	}

	pulseCnt += 1;

	if($5 > 0) {
		renderCnt += 1;
		nodes += $5;
		layoutPass += $6;
		cssPass += $7;
		updateBounds += $8;
		waiting += $9;
		copyState += $10;
		dirtyOpts += $11;
		renderRoot += $12;
		painting += $13;
	}

	lastRenderCnt = $5;
}

END {
	printAvgs();
}
