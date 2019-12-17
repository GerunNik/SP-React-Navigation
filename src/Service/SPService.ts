import { INavLink } from 'office-ui-fabric-react/lib/Nav';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

export class SPService {
  public static async GetAnchorLinks(context: WebPartContext) {
    let anchorLinks: INavLink[] = [];

    try {
      // /* Page ID on which the web part is added */
      // let pageId = context.pageContext.listItem.id;

      // /* Get the canvasContent1 data for the page which consists of all the HTML */
      // let data = await context.spHttpClient.get(`${context.pageContext.web.absoluteUrl}/_api/sitepages/pages(${pageId})`, SPHttpClient.configurations.v1);
      // let jsonData = await data.json();
      // let canvasContent1 = jsonData.CanvasContent1;
      // let canvasContent1JSON: any[] = JSON.parse(canvasContent1);

      // /* Initialize variables to be used for sorting and adding the Navigation links */
      let headingIndex = -1;
      let subHeadingIndex = -1;
      let subSubHeadingIndex = -1;
      // let subSubSubHeadingIndex = -1;
      let headingOrder = 0;
      let prevHeadingOrder = 0;

      // /* Array to store all unique anchor URLs */
      let allUrls: string[] = [];
      const canvasContent1JSON = [{innerHTML: '<div><h1>Erster Titel</h1><h2>Zweiter Titel</h2><h3>Dritter Titel</h3><h4>Vierter Titel</h4><h1>Zweiter erster Titel</h1><h2>Zweiter zweiter Titel</h2><h1>Dritter erster Titel</h1><h2>Dritter zweiter Titel</h2><h3>Dritter dritter Titel</h3></div>'}];

      /* Traverse through all the Text web parts in the page */
      canvasContent1JSON.map((webPart) => {
        if (webPart.innerHTML) {
          let HTMLString: string = webPart.innerHTML;

          while (HTMLString.search(/<h[1-4]>/g) !== -1) {
            /* The Header Text value */
            let headingValue = HTMLString.substring(HTMLString.search(/<h[1-4]>/g) + 4, HTMLString.search(/<\/h[1-4]>/g));
            console.log(headingValue);
            headingOrder = parseInt(HTMLString.charAt(HTMLString.search(/<h[1-4]>/g) + 2));

            /* Check if same anchorUrl already exists */
            let urlExists = true;
            let anchorUrl = `#${headingValue.replace(/ /g, '-')}`.toLowerCase();
            let urlSuffix = 1;
            while (urlExists === true) {
              urlExists = (allUrls.indexOf(anchorUrl) === -1) ? false : true;
              if (urlExists) {
                anchorUrl = anchorUrl + `-${urlSuffix}`;
                urlSuffix++;
              }
            }
            allUrls.push(anchorUrl);

            /* Add links to Nav element */
            if (anchorLinks.length === 0) {
              anchorLinks.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
              headingIndex++;
            } else {
              if (headingOrder <= prevHeadingOrder) {
                /* Adding or Promoting links */
                switch (headingOrder)
                {
                  case 1:
                      anchorLinks.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                      headingIndex++;
                    break;
                  case 2:
                      anchorLinks.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                      subHeadingIndex = -1;
                    break;
                  case 3:
                      anchorLinks[headingIndex].links.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                      subHeadingIndex++;
                    break;
                  case 4:
                      anchorLinks[headingIndex].links[subHeadingIndex].links.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                    break;
                }
              } else {
                /* Making sub links */
                if (headingOrder === 2) {
                  anchorLinks[headingIndex].links.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                  subHeadingIndex++;
                } 
                else if (headingOrder === 3) {
                  anchorLinks[headingIndex].links[subHeadingIndex].links.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                  subSubHeadingIndex++;
                } 
                else if (headingOrder === 4) {
                  anchorLinks[headingIndex].links[subHeadingIndex].links[subSubHeadingIndex].links.push({ name: headingValue, key: anchorUrl, url: anchorUrl, links: [], isExpanded: true });
                } 
              }
            }
            prevHeadingOrder = headingOrder;

            /* Replace the added header links from the string so they don't get processed again */
            HTMLString = HTMLString.replace(`<h${headingOrder}>`, '');
            HTMLString = HTMLString.replace(`</h${headingOrder}>`, '');
          }
        }
      });
    } catch (error) {
      console.log(error);
    }

    console.log(anchorLinks);
    return anchorLinks;
  }
}