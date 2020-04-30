import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as pdf from "pdfjs-dist";
import * as zip from "jszip";
import { saveAs } from "file-saver";
export class AttachmentView implements ComponentFramework.StandardControl<IInputs, IOutputs> {
	private _container: HTMLDivElement;
	private _context: ComponentFramework.Context<IInputs>;
	//下载数组
	private _doarr: string[] = [];
	//进度条
	private _ing: HTMLDivElement;
	//主div
	private _div: HTMLDivElement;
	//base64
	private _base64: string | null;
	//搜索框
	private _Google: HTMLInputElement;
	//页面指示
	private _lookpage: HTMLInputElement;
	//遮罩层
	private _mark: HTMLDivElement;
	//pdf画布
	private _canvas: HTMLCanvasElement;
	//pdf总页数
	private _pdfpages: number;
	//pdf当前页
	private _page: number;
	//图像类型
	private _png_type: string[] = ["png", "jpg","gif"];
	//预览图片
	private _img: HTMLImageElement;
	//音频类型
	private _db_type: string[] = ["mp3"];
	//预览音频
	private _db: HTMLAudioElement;
	//浏览器支持文件
	private _Browser_type: string[] = ["png", "jpg","mp3","pdf","gif","json","txt","xml"];
	constructor() { }

	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
		let HTML: Document = document;
		this._container = container;
		this._context = context;
		//自动宽度
		this._context.mode.trackContainerResize(true);
		//获取当前窗体记录Guid与实体名称
		// @ts-ignore
		let recordguid: string = this._context.mode.contextInfo.entityId;
		// @ts-ignore
		let recordname: string = this._context.mode.contextInfo.entityTypeName;
		//主div
		this._div = HTML.createElement("div");
		this._div.setAttribute("class", "div_css");
		//搜索框
		this._Google = HTML.createElement("input");
		this._Google.type = "text";
		this._Google.placeholder = "咕噜咕噜?";
		this._Google.title = "模糊搜索,区分大小写,通过|分割多个关键词";
		this._Google.setAttribute("class", "select");
		this._Google.addEventListener("blur", (x) => this.selectdiv(this._Google.value));
		//打包下载按钮
		let dow = HTML.createElement("input");
		dow.type = "button";
		dow.value = "下载";
		dow.setAttribute("class","dowbut");
		dow.addEventListener("click", (x) => this.download());
		//下载进度条
		this._ing = HTML.createElement("div");
		this._ing.style.height = "6px";
		this._ing.style.transition = "width 0.3s";
		this._ing.style.backgroundColor = "#B5495B";
		this._ing.style.width = "0px";
		this._ing.hidden = true;

		//页面指示
		this._lookpage = HTML.createElement("input");
		this._lookpage.readOnly = true;
		this._lookpage.type = "text";
		this._lookpage.setAttribute("class", "lookpage");
		//创建遮罩,默认隐藏
		if (!HTML.getElementById("mark")) {
			this._mark = HTML.createElement("div");
			this._mark.setAttribute("class", "mark");
			this._mark.setAttribute("id", "mark");
			this._mark.hidden = true;
			//创建遮罩事件
			this._mark.addEventListener("click", (x) => {
				if ((<HTMLDivElement>x.target!).id == "mark") {
					//隐藏遮罩,删除遮罩子元素
					this._mark.hidden = true;
					//this._lookpage.value = "加载Base64中";
					this._mark.innerHTML = "";
					this._base64 = null;
				}
			});
			//添加遮罩
			let boby = HTML.getElementsByTagName("body")[0];
			boby.appendChild(this._mark);
		}
		else {
			this._mark = <HTMLDivElement>HTML.getElementById("mark");
		}
		//查询记录相关附件名称

		let query = `?$select=filename&$filter=filename ne null and _objectid_value eq ${recordguid} and objecttypecode eq '${recordname}'`;
		this._context.webAPI.retrieveMultipleRecords("annotation", query).then((json) => {
			json.entities.forEach((key) => {
				let ol: HTMLOListElement = HTML.createElement("ol");
				ol.setAttribute("class", "ol_css");
				//名称
				ol.textContent = key["filename"];
				//Guid
				ol.id = key["annotationid"];
				//文件类型,转小写
				ol.setAttribute("_type", <string>key["filename"].split(".").pop()!.toLowerCase());
				//点击事件
				ol.addEventListener("click", (x) => this.iftype(ol.id, ol.getAttribute("_type")!));

				//下载按钮
				let but: HTMLInputElement = HTML.createElement("input");
				but.id = key["annotationid"];
				but.type = "button";
				but.setAttribute("class", "but");
				but.setAttribute("OVO", "false");
				but.addEventListener("click", (x) => {
					//获取触发事件的元素
					let button = <HTMLInputElement>x.target!
					if (button.getAttribute("OVO") == "false") {
						button.setAttribute("OVO", "true");
						but.style.backgroundColor = "#B5495B";
						//将要下载的Guid添加进下载数组
						this._doarr.push(button.id);
					} else {
						button.setAttribute("OVO", "false");
						but.style.backgroundColor = "#FFD32E";
						//从下载数组删除guid
						let index = this._doarr.indexOf(button.id);
						this._doarr.splice(index, 1);
					}
				})
				let div: HTMLDivElement = HTML.createElement("div");
				div.appendChild(but);
				div.appendChild(ol);
				this._div.appendChild(div);
			})
		},
			(error) => {
				console.log(error);
			}

		)
		
		this._container.appendChild(this._ing);
		let div :HTMLDivElement = HTML.createElement("div");
		div.appendChild(this._Google);
		div.appendChild(dow);
		this._container.appendChild(div);
		this._container.appendChild(this._div);
		pdf.GlobalWorkerOptions.workerSrc = this._context.parameters.PDFjsurl.raw!;
	}
	//判断文件类型
	public iftype(Guid: string, type: string): void {
		if (this._context.parameters.BrowserView.raw =="True"){
			//浏览器支持的文件则使用浏览器打开
			if (this._Browser_type.indexOf(type) != -1){
				this.BrowserView(Guid);
			}
			else if (type == "zip") {
				this.lookZip(Guid);
			}
			else {alert("不支持的文件")}
		}
			//使用内置库查看
		else{
			if (type == "pdf") {
				this.lookpdf(Guid);
			}
			else if (this._png_type.indexOf(type) != -1) {
				this.lookimg(Guid);
			}
			else if (type == "zip") {
				this.lookZip(Guid);
			}
			else if (this._db_type.indexOf(type) != -1) {
				this.lookdb(Guid);
			}
			//没有内置库支持的文件尝试浏览器打开
			else{
				if (this._Browser_type.indexOf(type) != -1){
					this.BrowserView(Guid);
				}
				else{
					alert("不支持的文件")
				}
			}
		}
	}
	//显示图片预览
	public lookimg(Guid: string): void {
		//根据Guid查询注释实体
		this._context.webAPI.retrieveRecord("annotation", Guid, "?$select=documentbody,mimetype").then(
			//成功时
			(json) => {
				this._base64 = "data:" + json["mimetype"] + ";base64," + json["documentbody"];
				//显示遮罩
				this._mark.hidden = false;
				//创建Img
				this._img = document.createElement("img");
				this._img.src = this._base64;
				this._img.setAttribute("class", "img_css");
				this._mark.appendChild(this._img);
			},
			(error) => { console.log(error); alert("查找base64失败,请检查是否存在:" + Guid) }
		);
	}
	//显示PDF预览
	public lookpdf(Guid: string): void {
		//显示遮罩
		this._mark.hidden = false;
		this._lookpage.value = "加载Base64中";
		this._mark.appendChild(this._lookpage);
		//根据Guid查询注释实体
		this._context.webAPI.retrieveRecord("annotation", Guid, "?$select=documentbody").then(
			//成功时
			(json) => {
				this._base64 = window.atob(json["documentbody"]);
				//创建pdf画布
				this._canvas = document.createElement("canvas");
				this._canvas.setAttribute("class", "pdf_css");
				this._mark.appendChild(this._canvas);
				//初始化pdf页面
				let pdfdoc = pdf.getDocument({ data: this._base64 });
				//加载pdf,初始第一页
				this._page = 1;
				this.renderPage(pdfdoc, this._page);
				//创建页面切换按钮
				let next: HTMLInputElement = document.createElement("input");
				next.type = "button";
				next.id = "pdf_next";
				next.value = ">";
				next.setAttribute("class", "page_next_css");
				next.addEventListener("click", (x) => this.Switchpage((<HTMLDivElement>x.target!).id, pdfdoc));
				let upper: HTMLInputElement = document.createElement("input");
				upper.type = "button";
				upper.id = "pdf_upper";
				upper.value = "<";
				upper.setAttribute("class", "page_upper_css");
				upper.addEventListener("click", (x) => this.Switchpage((<HTMLDivElement>x.target!).id, pdfdoc));
				//添加按钮
				this._mark.appendChild(upper);
				this._mark.appendChild(next);
			},
			(error) => { console.log(error); alert("查找base64失败,请检查是否存在:" + Guid); this._lookpage.value = "下载base64失败!" }
		);
	}
	//判断pdf页面
	public Switchpage(e: string, pdfdoc: pdf.PDFLoadingTask<pdf.PDFDocumentProxy>): void {
		if (e == "pdf_next") {
			if (this._page == this._pdfpages) {
				alert("已到达尾页");
			}
			else {
				this._page = this._page + 1;
				this.renderPage(pdfdoc, this._page);
			}
		}
		else {
			if (this._page == 1) {
				alert("已到达第一页");
			}
			else {
				this._page = this._page - 1;
				this.renderPage(pdfdoc, this._page);
			}
		}
	}
	//加载PDF页面
	public renderPage(pdfdoc: pdf.PDFLoadingTask<pdf.PDFDocumentProxy>, _page: number): void {
		pdfdoc.promise.then((pdf) => {
			this._pdfpages = pdf.numPages;
			//加载页面
			pdf.getPage(_page).then((page) => {
				this._lookpage.value = `${_page}/${this._pdfpages}页`;
				//缩放等级100%
				let viewport: pdf.PDFPageViewport = page.getViewport({ scale: 1 });
				let context = this._canvas.getContext('2d');
				this._canvas.height = viewport.height;
				this._canvas.width = viewport.width;
				let renderContext = {
					canvasContext: context!,
					viewport: viewport
				};
				let renderTask = page.render(renderContext);
				renderTask.promise.then(function () {
				});
			}, (error) => { console.log(`加载pdf第${_page}页失败: ` + error); this._lookpage.value = `加载第${_page}页失败`; }
			)
		})
	}
	//预览ZIP文件内容
	public lookZip(Guid: string): void {
		//显示遮罩
		this._mark.hidden = false;
		this._lookpage.value = "加载Base64中";
		this._mark.appendChild(this._lookpage);
		//根据Guid查询注释实体
		this._context.webAPI.retrieveRecord("annotation", Guid, "?$select=documentbody").then(
			//成功时
			(json) => {
				this._base64 = <string>json["documentbody"];
				//创建显示zip内文件与文件数量列表
				let zipdiv: HTMLDivElement = document.createElement("div");
				zipdiv.style.height = "100%";
				zipdiv.style.overflow = "auto";
				zipdiv.style.position = "fixed";
				zipdiv.style.left = "50%";
				zipdiv.setAttribute("class", "div_css")
				//zip内文件计数
				let count: number = 0;
				//打开zip文件,默认UTF-8解码
				zip.loadAsync(this._base64, { base64: true }).then((zip) => {
					zip.forEach((name) => {
						//只显示文件
						if (name[name.length - 1] != "/") {
							count = count + 1
							let ol: HTMLOListElement = document.createElement("ol");
							ol.setAttribute("class", "ol_css");
							//名称
							ol.textContent = name.split("/").pop()!.toLowerCase();
							//路径
							ol.title = name;
							let div: HTMLDivElement = document.createElement("div");
							div.appendChild(ol);
							zipdiv.appendChild(div);
						}
					})
					this._lookpage.value = `共${count}个项目`
				},
					(error) => { console.log("打开zip文件失败:" + error); this._lookpage.value = "打开zip文件失败" }
				);
				this._mark.appendChild(zipdiv);
			},
			(error) => { console.log(error); alert("查找base64失败,请检查是否存在:" + Guid); this._lookpage.value = "下载base64失败!" }
		);
	}
	//筛选列表
	public selectdiv(value: string): void {
		//获取div列表
		let divlist: HTMLCollection = this._div.children;
		//判断多项筛选
		let arr: string[] = value.indexOf("|") != -1 ? value.split("|") : [value];
		for (let index = 0; index < divlist.length; index++) {
			//获取div
			let record: HTMLDivElement = <HTMLDivElement>divlist[index]
			let recordvalue: string = record.innerText;

			if (arr.filter((x) => recordvalue.indexOf(x) != -1).length == 0) {
				record.style.display = "none"
			}
			else {
				record.style.display = ""
			}
		}

	}
	//预览音频
	public lookdb(Guid: string): void {
		//设置高斯模糊
		this._div.style.filter = "blur(2.5px)";
		this._Google.style.filter = "blur(2.5px)";
		//根据Guid查询注释实体
		this._context.webAPI.retrieveRecord("annotation", Guid, "?$select=documentbody,mimetype").then(
			//成功时
			(json) => {
				this._base64 = "data:" + json["mimetype"] + ";base64," + json["documentbody"];
				//设置关闭按钮
				let X: HTMLInputElement = document.createElement("input");
				X.type = "button";
				X.setAttribute("id", "NTR");
				X.setAttribute("class", "X");
				X.value = "X"
				X.addEventListener("click", () => {
					this._div.style.filter = "";
					this._Google.style.filter = "";
					this._db.remove();
					X.remove();
				});
				this._container.appendChild(X);
				//创建db
				this._db = document.createElement("audio");
				this._db.src = this._base64;
				this._db.controls = true;
				this._db.setAttribute("class", "db_css");
				this._container.appendChild(this._db);
			},
			(error) => {
				console.log(error);
				alert("查找base64失败,请检查是否存在:" + Guid);
				this._div.style.filter = "";
				this._Google.style.filter = "";
			}
		);
	}
	//下载文件
	public async download() {
		let count = this._doarr.length;
		if (count <= 0) {
			alert("没有选择一个要下载的文件呢?");
		}
		else {
			//显示进度条
			this._ing.hidden = false;
			//创建zip
			let Zip = new zip
			for (let index = 0; index < count; index++) {
				//控制进度条宽度
				this._ing.style.width = ((index + 1) / count * 100).toString() + "%";
				//等待异步请求完全加载
				await this._context.webAPI.retrieveRecord("annotation", this._doarr[index], "?$select=documentbody,filename").then(
					//成功时
					(json) => {
						//添加文件						
						Zip.file(json["filename"], json["documentbody"], { base64: true });
					}
					//失败时
					,(error)=>{console.log(error);alert("文件:"+this._doarr[index]+"下载失败,无法找到注释记录")}
				)
				if (index+1 == count){this.OK(Zip)}
			}
		}

	}
	//浏览器保存
	public OK(Zip: zip): void {
		Zip.generateAsync({ type: "blob"}).then(
			(context) => {
				saveAs(context, "附件.zip");
				this._ing.hidden = true;
				this._ing.style.width = "0px";
			},
			(err) => { alert("创建zip文件时出错"); console.log(err) }
		)
	}
	//使用浏览器预览
	public async BrowserView(Guid:string){
		//显示进度条
		this._ing.hidden = false;
		this._ing.style.width="20%"
		await this._context.webAPI.retrieveRecord("annotation", Guid, "?$select=documentbody,mimetype").then(
			//成功时
			(json) => {
					//解码
					var byteString = atob(json["documentbody"]); 
					var arrayBuffer = new ArrayBuffer(byteString.length);
					var intArray = new Uint8Array(arrayBuffer);
					
					for (var i = 0; i < byteString.length; i++) {
						intArray[i] = byteString.charCodeAt(i);
					}
					let blob = new Blob([intArray], {type: json["mimetype"]});
					//打开标签页
					window.open(URL.createObjectURL(blob));
			},
			(error)=>{console.log(error); alert("查找base64失败,请检查是否存在:" + Guid)}
		)
		this._ing.hidden = true;
		this._ing.style.width = "0px";
	}
	public updateView(context: ComponentFramework.Context<IInputs>): void {
	}
	public getOutputs(): IOutputs { return {}; }
	public destroy(): void { }
}