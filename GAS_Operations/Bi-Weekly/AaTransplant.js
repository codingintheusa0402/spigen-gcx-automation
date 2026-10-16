// TEMP (bi-weekly-builder): move this period's finished Apple slides into the golden-base copy. Delete after running.
function aaTransplant() {
  const dst = SlidesApp.openById('1sJ49LsEqMFrFyykAsgpWIREO9SCnWa5GW-ILEBLbG8Q'), src = SlidesApp.openById('1F_zb9N_R0oAxBKCmDeEUxqSBY-demuW7Woldh2Ok4SQ');
  const S = src.getSlides(), D = dst.getSlides();
  // dst keeps: 0 cover, 1 index, 2 dashboard, 76 appendix, 80 closing → remove the rest
  D.forEach(function(sl, i) { if ([0, 1, 2, 76, 80].indexOf(i) < 0) sl.remove(); });
  // now dst = [cover, index, dashboard, appendix, closing]
  let at = 3;
  for (let i = 3; i <= 54; i++) dst.insertSlide(at++, S[i]);          // 전제품, claim TOP3, TOP-7, cards, SIREN
  at++;                                                                 // skip appendix
  for (let i = 56; i <= 58; i++) dst.insertSlide(at++, S[i]);          // bad-review TOP3 backups
  // weekly chart: swap the golden chart image for this period's
  const srcChart = S[2].getImages().filter(function(im) { return im.getTitle() === 'apl_chart'; })[0];
  const dstChart = dst.getSlides()[2].getImages().filter(function(im) { return im.getTitle() === 'apl_chart'; })[0];
  if (srcChart && dstChart) dstChart.replace(srcChart.getBlob());
  dst.saveAndClose();
  Logger.log('slides now ' + SlidesApp.openById('1sJ49LsEqMFrFyykAsgpWIREO9SCnWa5GW-ILEBLbG8Q').getSlides().length);
}
