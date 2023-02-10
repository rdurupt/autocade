var laTailleMax=0;
var NTrees=[];
//Detection Navigateur
function bw_check()
{
  var is_major = parseInt(navigator.appVersion);
  this.ver=navigator.appVersion;
  this.agent=navigator.userAgent;
  this.dom=document.getElementById?1:0;
  this.opera=this.agent.indexOf("Opera")>-1;
  this.ie5=(this.ver.indexOf("MSIE 5")>-1 && this.dom && !this.opera)?1:0;
  this.ie6=(this.ver.indexOf("MSIE 6")>-1 && this.dom && !this.opera)?1:0;
  this.ie4=(document.all && !this.dom && !this.opera)?1:0;
  this.ie=this.ie4||this.ie5||this.ie6;
  this.mac=this.agent.indexOf("Mac")>-1;
  this.ns6=(this.dom && parseInt(this.ver) >= 5)?1:0;
  this.ie3=(this.ver.indexOf("MSIE") && (is_major < 4));
  this.hotjava=(this.agent.toLowerCase().indexOf('hotjava') != -1)?1:0;
  this.ns4=(document.layers && !this.dom && !this.hotjava)?1:0;
  this.bw=(this.ie6 || this.ie5 || this.ie4 || this.ns4 || this.ns6 || this.opera);
  this.ver3=(this.hotjava || this.ie3);
  return this;
};
//Initialisation des options
function Portal_Format( fmt, tree )
{
  this.init = function( fmt, tree )
  {
    this.left=fmt[0];
    this.top=fmt[1];
    this.showB=fmt[2];
    this.nP=fmt[3][0];
    this.nM=fmt[3][1];
    this.nLP=fmt[3][2];
    this.nLM=fmt[3][3];
    this.nL=fmt[3][4];
    this.nT=fmt[3][5];
    this.nV=fmt[3][6];
    this.nB=fmt[3][7];
    this.Bw=fmt[4][0];
    this.Bh=fmt[4][1];
    this.Ew=fmt[4][2];
    this.showF=fmt[5];
    this.clF=fmt[6][0];
    this.exF=fmt[6][1];
    this.iF=fmt[6][2];
    this.Fw=fmt[7][0];
    this.Fh=fmt[7][1];
    this.ident=fmt[8];
    this.back=new Portal_Back(this.left, this.top, fmt[9], 'cls'+tree.name+'_back');
    this.nst=fmt[10];
    this.nstl=fmt[11];
    this.so=fmt[12];
    this.pg=fmt[13][0];
    this.sp=fmt[13][1];
	this.aLink=fmt[14];
	this.showBACK=fmt[15];
	if(this.showBACK){
	  this.bgG=fmt[16][0];
	  this.bgF=fmt[16][1];
	  this.bgD=fmt[16][2];
	  this.bgW=fmt[17][0];
	  this.bgH=fmt[17][1];
	}
	this.showTOP=fmt[18];
	if(this.showTOP){
	  this.topG=fmt[19][0];
	  this.topF=fmt[19][1];
	  this.topD=fmt[19][2];
	  this.topW=fmt[20][0];
	  this.topH=fmt[20][1];
	}
	this.showBOT=fmt[21];
	if(this.showBOT){
	  this.botG=fmt[22][0];
	  this.botF=fmt[22][1];
	  this.botD=fmt[22][2];
	  this.botW=fmt[23][0];
	  this.botH=fmt[23][1];
	}
    if(this.showB){
      this.e=new Image();
      this.e.src=this.nP;
      this.e1=new Image();
      this.e1.src=this.nM;
      this.e2=new Image();
      this.e2.src=this.nLP;
      this.e3=new Image();
      this.e3.src=this.nLM;
      this.e4=new Image();
      this.e4.src=this.nL;
      this.e5=new Image();
      this.e5.src=this.nT;
      this.e6=new Image();
      this.e6.src=this.nV;
      this.e7=new Image();
      this.e7.src=this.nB;
    }
    if(this.showF){
      this.e8=new Image();
      this.e8.src=this.exF;
      this.e9=new Image();
      this.e9.src=this.clF;
      this.e10=new Image();
      this.e10.src=this.iF;
    }
    if(this.showBACK){
      this.e11=new Image();
      this.e11.src=this.bgG;
      this.e12=new Image();
      this.e12.src=this.bgF;
      this.e13=new Image();
      this.e13.src=this.bgD;
    }
    if(this.showTOP){
      this.e14=new Image();
      this.e14.src=this.topG;
      this.e15=new Image();
      this.e15.src=this.topF;
      this.e16=new Image();
      this.e16.src=this.topD;
    }
    if(this.showBOT){
      this.e17=new Image();
      this.e17.src=this.botG;
      this.e18=new Image();
      this.e18.src=this.botF;
      this.e19=new Image();
      this.e19.src=this.botD;
    }
  };
  this.nstyle = function ( lvl )
  {
    return (und(this.nstl[lvl]))?this.nst:this.nstl[lvl];
  };
  this.idn = function( lvl )
  {
    var r=(und(this.ident[lvl]))?this.ident[0]*lvl:this.ident[lvl];
    return r;
  };
  this.init(fmt, tree);
};
//Construction calque background
function Portal_Back( inLeft, inTop, inColor, inName )
{
  this.bw=new bw_check();
  this.ns4=this.bw.ns4;
  this.left=inLeft;
  this.top=inTop;
  this.name=inName;
  this.color=inColor;
  this.resize = function( w, h )
  {
    if(this.ns4){
      this.el.resizeTo(w, h);
    }else{
      this.el.style.width=w;
      this.el.style.height=h;
      if(this.r){
        this.el2.style.top=h+this.top-5;
      }
    }
  };
  this.init = function()
  {
    if(this.r){
      if(!this.ns4){
        var bgc=this.color==""?"":" background-color:"+this.color+";";
        document.write('<div id="'+this.name+'c" style="'+bgc+'position:absolute;z-index:-1;top:'+this.top+'px;left:'+this.left+'px"></div>');
        this.el2=document.all?document.all[this.name+'c']:document.getElementById(this.name+'c');
      }
      if(this.ns4){
        var bgc=this.color==""?"":' bgcolor="'+this.color+'" ';
        document.write('<layer '+bgc+' top="'+this.top+'" left="'+this.left+'" id="'+this.name+'" z-index="0">'+ '</layer>');
        this.el=document.layers[this.name];
      }else{
        var bgc=this.color==""?"":" background-color:"+this.color+";";
        document.write('<div id="'+this.name+'" style="'+bgc+'position:absolute;z-index:0;top:'+this.top+'px;left:'+this.left+'px"></div>');
        this.el=document.all?document.all[this.name]:document.getElementById(this.name);
      }
    }
  };
  this.r=true;
  this.init();
};
//Construction de l'arbre
function Portal_Tree( name, nodes, format )
{
  this.bw=new bw_check();
  this.ns4=this.bw.ns4;
  this.name=name;
  this.fmt=new Portal_Format(format, this);
  if(typeof(NTrees)=='undefined'){
    NTrees=new Array();
  }
  NTrees[this.name]=this;
  this.Nodes=new Array();
  this.rootNode=new Portal_Node(null, "", "", "", null, "");
  this.rootNode.treeView=this;
  if(this.fmt.showBACK){
    this.NodesBACK=new Array();
    this.rootNodeBACK=new Portal_Node_Back(null, "", "");
    this.rootNodeBACK.treeView=this;
  }
  this.selectedNode=null;
  this.maxWidth = 0;
  this.maxHeight = 0;
  this.ondraw = null;	
  this.addNode = function ( node )
  {
    var parentNode=node.parentNode;
    this.Nodes=this.Nodes.concat([node]);
    node.index=this.Nodes.length-1;
    if(parentNode==null){
      this.rootNode.children=this.rootNode.children.concat([node]);
    }else{
      parentNode.children=parentNode.children.concat([node]);
    }  
    return node;
  };
  this.addNodeBACK = function ( node )
  {
    var parentNode=node.parentNode;
    this.NodesBACK=this.NodesBACK.concat([node]);
    node.index=this.NodesBACK.length-1;
    if(parentNode==null){
      this.rootNodeBACK.children=this.rootNodeBACK.children.concat([node]);
    }else{
      parentNode.children=parentNode.children.concat([node]);
    }  
    return node;
  };
  this.rebuildTree = function()
  {
    var s = "";
    for(var i=0; i<this.Nodes.length; i++){
      s+=this.Nodes[i].init();
    };
    document.write(s);
    for(var i=0; i<this.Nodes.length; i++){
      if(this.ns4){
        this.Nodes[i].el=document.layers[this.Nodes[i].id()+"d"];
        if(this.Nodes[i].getW()>laTailleMax){
          laTailleMax=this.Nodes[i].getW();
        }
        if(this.fmt.showF){
          this.Nodes[i].nf=this.Nodes[i].el.document.images[this.Nodes[i].id()+"nf"];
        }
        if(this.fmt.showB){
          this.Nodes[i].nb=this.Nodes[i].el.document.images[this.Nodes[i].id()+"nb"];
		}
      }else{
        this.Nodes[i].el=document.all?document.all[this.Nodes[i].id()+"d"]:document.getElementById(this.Nodes[i].id()+"d");
        if(this.Nodes[i].getW()>laTailleMax){
          laTailleMax=this.Nodes[i].getW();
        }
        if(this.fmt.showB){
          this.Nodes[i].nb=document.all?document.all[this.Nodes[i].id()+"nb"]:document.getElementById(this.Nodes[i].id()+"nb");
        }
        if(this.fmt.showF){
          this.Nodes[i].nf=document.all?document.all[this.Nodes[i].id()+"nf"]:document.getElementById(this.Nodes[i].id()+"nf");
        }
      }
	};
  };
  this.rebuildTreeBACK = function()
  {
    var s = "";
    for(var i=0; i<this.NodesBACK.length; i++){
      s+=this.NodesBACK[i].init();
    };
    document.write(s);
    for(var i=0; i<this.NodesBACK.length; i++){
      if(this.ns4){
        this.NodesBACK[i].el=document.layers[this.NodesBACK[i].id()+"b"];
      }else{
        this.NodesBACK[i].el=document.all?document.all[this.NodesBACK[i].id()+"b"]:document.getElementById(this.NodesBACK[i].id()+"b");
      }
	};
  };
  this.draw = function()
  {
    this.currTop=this.fmt.top;
    this.maxHeight=0;
    this.maxWidth=0;
	for(var i=0; i<this.rootNode.children.length; i++){
	  this.rootNode.children[i].draw(true);
	};
    this.fmt.back.resize(this.maxWidth-this.fmt.left, this.maxHeight-this.fmt.top);
    if(this.ondraw!=null){
      this.ondraw();
    }
  };
  this.drawBACK = function()
  {
    this.currTop=this.fmt.top;
    this.maxHeight=0;
    this.maxWidth=0;
	for(var i=0; i<this.rootNodeBACK.children.length; i++){
	  this.rootNodeBACK.children[i].draw(true);
	};
    this.fmt.back.resize(this.maxWidth-this.fmt.left, this.maxHeight-this.fmt.top);
    if(this.ondraw!=null){
      this.ondraw();
    }
  };
  this.updateImages = function ( node )
  {
    if(node.expanded){
      if(node.parentNode!=null){
        if(node.index==node.parentNode.children[node.parentNode.children.length-1].index){
          srcB=node.treeView.fmt.nLM;
        }else{
          srcB=node.treeView.fmt.nM;
        }
      }else{
        if(node.index==node.treeView.rootNode.children[node.treeView.rootNode.children.length-1].index){
          srcB=node.treeView.fmt.nLM;
        }else{
          srcB=node.treeView.fmt.nM;
        }
      }
    }else{
      if(node.parentNode!=null){
        if(node.index==node.parentNode.children[node.parentNode.children.length-1].index){
          srcB=node.treeView.fmt.nLP;
        }else{
          srcB=node.treeView.fmt.nP;
        }
      }else{
        if(node.index==node.treeView.rootNode.children[node.treeView.rootNode.children.length-1].index){
          srcB=node.treeView.fmt.nLP;
        }else{
          srcB=node.treeView.fmt.nP;
        }
      }
    }
    var srcF=node.expanded?this.fmt.exF:this.fmt.clF;
    if(node.treeView.fmt.showB && node.nb && node.nb.src != srcB){
      node.nb.src=srcB;
    }
    if(node.treeView.fmt.showF && node.nf && node.nf.src != srcF){
      node.nf.src=node.hasChildren()?srcF:this.fmt.iF;
    }
  };
  this.expandNode = function( index )
  {
    var node=this.Nodes[index];
    var pNode=node.parentNode?node.parentNode:null;
    if(this.fmt.showBACK){
      if(this.fmt.showTOP){
        var nodeBACK=this.NodesBACK[index+1];
      }else{
        var nodeBACK=this.NodesBACK[index];
      }
      var pNodeBACK=nodeBACK.parentNode?nodeBACK.parentNode:null;
    }
    if(!und(node) && node.hasChildren()){
      node.expanded=!node.expanded;
      if(this.fmt.showBACK){
        nodeBACK.expanded=!nodeBACK.expanded;
      }
      this.updateImages(node);
      if(!node.expanded){
        node.hideChildren();
        if(this.fmt.showBACK){
          nodeBACK.hideChildren();
        }
      }else{
        if (this.fmt.so){
          for(var i=0; i<this.Nodes.length; i++){
            this.Nodes[i].show(false);
            if(this.fmt.showBACK){
              if(this.fmt.showTOP){
                this.NodesBACK[i+1].show(false);
              }else{
                this.NodesBACK[i].show(false);
              }
            }
            if(this.Nodes[i]!=node && this.Nodes[i].parentNode==pNode){
              this.Nodes[i].expanded=false;
              if(this.fmt.showBACK){
                if(this.fmt.showTOP){
                  this.NodesBACK[i+1].expanded=false;
                }else{
                  this.NodesBACK[i].expanded=false;
                }
              }
              this.updateImages(this.Nodes[i]);
            }
          };
        }
      }
      this.draw();
      if(this.fmt.showBACK){
        this.drawBACK();
      }
    }
  };
  this.selectNode = function( index )
  {
    var node=this.Nodes[index];
    if(!und(node)){
      this.selectedNode = node;
      node.draw();
    }
  };
  this.readNodes = function ( nodes )
  {
    var ind=0;
    var par=null;
    function readOne( arr , tree )
    {
      if(und(arr)){
        return;
      }
      var text=arr[0];
      var url=arr[1]==null?"":arr[1];
      var targ=arr[2]==null?"_self":arr[2];
	  var cssClass=arr[3]==null?"":arr[3];
      var node=tree.addNode(new Portal_Node(tree, par, text, url, targ, cssClass))
      var i=4;
      while(!und(arr[i])){
        par=node;
        readOne(arr[i], tree);
        i++;
      };
    };
    if(und(nodes) || und(nodes[0]) || und(nodes[0][0])){
      return;
    }
    for(var i=0; i<nodes.length; i++){
      par=null;
      readOne(nodes[i], this);
    };
  };
  this.readNodesBACK = function ( nodes )
  {
    var ind=0;
    var par=null;
    var node=null;
    function readOneBACK( arr , tree , node )
    {
      if(und(arr)){
        return;
      }
      node=tree.addNodeBACK(new Portal_Node_Back(tree, par, "BACK"));
      var i=4;
      while(!und(arr[i])){
        par=node;
        readOneBACK(arr[i], tree, node);
        i++;
      };
    };
    if(und(nodes) || und(nodes[0]) || und(nodes[0][0])){
      return;
    }
    if(this.fmt.showTOP){
      node=this.addNodeBACK(new Portal_Node_Back(this, par, "TOP"));
    }
    for(var i=0; i<nodes.length; i++){
      par=null;
      readOneBACK(nodes[i], this, node);
    };
    if(this.fmt.showBOT){
      node=this.addNodeBACK(new Portal_Node_Back(this, par, "BOT"));
    }
  };
  this.init = function()
  {
    this.readNodes(nodes);
    this.rebuildTree();
    this.draw();
    if(this.fmt.showBACK){
      this.readNodesBACK(nodes);
      this.rebuildTreeBACK();
      this.drawBACK();
    }
  };
  this.init();
};
//Construction Node
function Portal_Node( treeView, parentNode , text, url, target, cssClass)
{
  this.index=-1;
  this.treeView=treeView;
  this.parentNode=parentNode;
  this.text=text;
  this.url=url;
  this.target=target;
  this.cssClass=cssClass;
  this.expanded=false;
  this.children=new Array();
  this.level = function()
  {
    var node=this;
    var i=0;
    while(node.parentNode!=null){
      i++;
      node=node.parentNode;
    };
    return i;
  };
  this.hasChildren = function()
  {
    return this.children.length>0;
  };
  this.init = function()
  {
    var s = "";
    if(this.treeView.ns4){
      s = '<layer id="'+this.id()+'d" z-index="'+this.index+10+'" visibility="hidden">'+this.getContent()+'</layer>';
    }else{
      s = '<div id="'+this.id()+'d" style="position:absolute;visibility:hidden;z-index:'+this.index+10+';">'+this.getContent()+'</div>';
    }
    return s;
  };
  this.getH = function()
  {
    return this.treeView.ns4?this.el.clip.height:this.el.offsetHeight;
  };
  this.getW = function()
  {
    return this.treeView.ns4?this.el.clip.width:this.el.offsetWidth;
  };
  this.id = function()
  {
    return 'nt'+this.treeView.name+this.index;
  };
  this.getContent = function()
  {
    function itemSquare(node){
      var img=node.hasChildren()?(node.expanded?node.treeView.fmt.exF:node.treeView.fmt.clF) : node.treeView.fmt.iF;
      var w=node.treeView.fmt.Fw;
      var h=node.treeView.fmt.Fh;
      return "<td valign=\"middle\" width=\""+w+"\"><img id=\""+node.id()+"nf\" name=\""+node.id()+"nf\" src=\"" + img + "\" width="+w+" height="+h+" border=0></td>";
    };
    function buttonSquare(node){
      if(node.hasChildren()){
        if(node.expanded){
          if(node.parentNode!=null){
            if(node.index==node.parentNode.children[node.parentNode.children.length-1].index){
              img=node.treeView.fmt.nLM;
            }else{
              img=node.treeView.fmt.nM;
            }
          }else{
            if(node.index==node.treeView.rootNode.children[node.treeView.rootNode.children.length-1].index){
              img=node.treeView.fmt.nLM;
            }else{
              img=node.treeView.fmt.nM;
            }
          }
        }else{
          if(node.parentNode!=null){
            if(node.index==node.parentNode.children[node.parentNode.children.length-1].index){
              img=node.treeView.fmt.nLP;
            }else{
              img=node.treeView.fmt.nP;
            }
          }else{
            if(node.index==node.treeView.rootNode.children[node.treeView.rootNode.children.length-1].index){
              img=node.treeView.fmt.nLP;
            }else{
              img=node.treeView.fmt.nP;
            }
          }
        }
      }else{
        if(node.parentNode!=null){
          if(node.index==node.parentNode.children[node.parentNode.children.length-1].index){
            img=node.treeView.fmt.nL;
          }else{
            img=node.treeView.fmt.nT;
          }
        }else{
          if(node.index==node.treeView.rootNode.children[node.treeView.rootNode.children.length-1].index){
            img=node.treeView.fmt.nL;
          }else{
            img=node.treeView.fmt.nT;
          }
        }
      }
      var w=node.treeView.fmt.Bw;
      var h=node.treeView.fmt.Bh;
      return '<td valign=\"middle\" width="'+w+'"><a href="javascript:NTrees[\''+node.treeView.name+'\'].expandNode('+node.index+')"><img name=\''+node.id()+'nb\' id=\''+node.id()+'nb\' src="' + img + '" width="'+w+'" height="'+h+'" border=0></a></td>';
    };
    function blankSquare(node, ww){
	  t="";
	  uneNode = node;
      if(uneNode.level()>1){
        while(uneNode.level()>1){
          if(uneNode.parentNode.parentNode.children.length>1){
            if(uneNode.parentNode.index!=uneNode.parentNode.parentNode.children[uneNode.parentNode.parentNode.children.length-1].index){
              t='<img src="'+node.treeView.fmt.nV+'" width="'+node.treeView.fmt.Bw+'" height="'+node.treeView.fmt.Bh+'" border="0">'+t;
            }else{
              t='<img src="'+node.treeView.fmt.nB+'" width="'+node.treeView.fmt.Bw+'" height="'+node.treeView.fmt.Bh+'" border="0">'+t;
            }
          }else{
            t='<img src="'+node.treeView.fmt.nB+'" width="'+node.treeView.fmt.Bw+'" height="'+node.treeView.fmt.Bh+'" border="0">'+t;
          }
          uneNode=uneNode.parentNode;
        };
        if(uneNode.parentNode.index!=node.treeView.rootNode.children[node.treeView.rootNode.children.length-1].index){
          t='<img src="'+node.treeView.fmt.nV+'" width="'+node.treeView.fmt.Bw+'" height="'+node.treeView.fmt.Bh+'" border="0">'+t;
        }else{
          t='<img src="'+node.treeView.fmt.nB+'" width="'+node.treeView.fmt.Bw+'" height="'+node.treeView.fmt.Bh+'" border="0">'+t;
        }
      }else{
        if(uneNode.parentNode.index!=node.treeView.rootNode.children[node.treeView.rootNode.children.length-1].index){
          t='<img src="'+node.treeView.fmt.nV+'" width="'+node.treeView.fmt.Bw+'" height="'+node.treeView.fmt.Bh+'" border="0">'+t;
        }else{
          t='<img src="'+node.treeView.fmt.nB+'" width="'+node.treeView.fmt.Bw+'" height="'+node.treeView.fmt.Bh+'" border="0">'+t;
        }
      }
      return "<td width=\""+ww+"\">"+t+"</td>"
    };
    var s = '';
    var ll = this.level();
    s+='<table cellpadding='+this.treeView.fmt.pg+' cellspacing='+this.treeView.fmt.sp+' border=0><tr>';
    var idn = this.treeView.fmt.idn(ll);
    if(idn>0){
      s+=blankSquare(this, idn);
    }
    if(this.treeView.fmt.showB){
      s+=buttonSquare(this);
    }
    if(this.treeView.fmt.showF){
      s+=itemSquare(this);
    }
    this.onStatus=this.text;
    while((this.onStatus).indexOf("'") > -1){
      this.onStatus=(this.onStatus).replace("'","\'");
    };
    if(this.url==""){
      if(this.treeView.fmt.aLink){
        s+=this.hasChildren()?'<td nowrap=\"1\" valign=\"top\">&nbsp;<a href="javascript:NTrees[\''+this.treeView.name+'\'].expandNode('+this.index+')" onMouseOver="window.status=\'' + this.onStatus + '\';return true;" onMouseOut="window.status=\' \';return true;"><font class="'+(this.cssClass!=""?this.cssClass:this.treeView.fmt.nstyle(ll))+'">'+this.text+'</font></a></td></tr></table>':'<td nowrap=\"1\" valign=\"top\"><a href="#" onMouseOver="window.status=\'' + this.onStatus + '\';return true;" onMouseOut="window.status=\' \';return true;"><font class="'+(this.cssClass!=""?this.cssClass:this.treeView.fmt.nstyle(ll))+'">'+this.text+'</font></a></td></tr></table>';
      }else{
        s+=this.hasChildren()?'<td nowrap=\"1\" valign=\"top\">&nbsp;<font class="'+(this.cssClass!=""?this.cssClass:this.treeView.fmt.nstyle(ll))+'">'+this.text+'</font></td></tr></table>':'<td nowrap=\"1\"><font class="'+(this.cssClass!=""?this.cssClass:this.treeView.fmt.nstyle(ll))+'">'+this.text+'</font></td></tr></table>';
      }
    }else{
      if(this.treeView.fmt.aLink){
        s+='<td nowrap=\"1\" valign=\"top\">&nbsp;<a href="'+this.url+'" target="'+this.target+'" onclick="javascript:NTrees[\''+this.treeView.name+'\'].expandNode('+this.index+')" onMouseOver="window.status=\'' + this.onStatus + '\';return true;" onMouseOut="window.status=\' \';return true;"><font class="'+(this.cssClass!=""?this.cssClass:this.treeView.fmt.nstyle(ll))+'">'+this.text+'</font></a></td></tr></table>';
      }else{
        s+='<td nowrap=\"1\" valign=\"top\">&nbsp;<a href="'+this.url+'" target="'+this.target+'" onMouseOver="window.status=\'' + this.onStatus + '\';return true;" onMouseOut="window.status=\' \';return true;"><font class="'+(this.cssClass!=""?this.cssClass:this.treeView.fmt.nstyle(ll))+'">'+this.text+'</font></a></td></tr></table>';
      }
    }
    return s;
  };
  this.moveTo = function( x, y )
  {
    if(this.treeView.ns4){
      this.el.moveTo(x,y);
    }else{
      this.el.style.left=x;
      this.el.style.top=y;
    }
  };
  this.show = function( sh )
  {
    if(this.visible==sh){
      return;
    }
    this.visible=sh;
    var vis=this.treeView.ns4?(sh?'show':'hide'):(sh?'visible':'hidden');
    if(this.treeView.ns4){
      this.el.visibility=vis;
    }else{
      this.el.style.visibility=vis;
    }
  };
  this.hideChildren = function()
  {
    this.show(false);
    for(var i=0; i<this.children.length; i++){
      this.children[i].hideChildren();
    };
  };
  this.draw = function()
  {
    var ll=this.treeView.fmt.left;
    if(this.treeView.fmt.showBACK){
      ll=ll+(this.treeView.fmt.bgW/2);
    }
    var tt=this.treeView.currTop;
    if(this.treeView.fmt.showTOP&&this.treeView.fmt.showBACK){
      tt=tt+this.treeView.fmt.topH;
    }
    this.moveTo(ll, tt);
    if(ll+this.getW()>this.treeView.maxWidth){
      this.treeView.maxWidth=ll+this.getW();
    }
    this.show(true);
    this.treeView.currTop+=this.getH();
    if(this.treeView.currTop>this.treeView.maxHeight){
      this.treeView.maxHeight=this.treeView.currTop;
    }
    if(this.expanded && this.hasChildren()){
      for(var i=0; i<this.children.length; i++){
        this.children[i].draw();
      };
    }
  };
};
//Construction Background Image
function Portal_Node_Back( treeView, parentNode, leCase )
{
  this.index=-1;
  this.treeView=treeView;
  this.leCase=leCase;
  this.parentNode=parentNode;
  this.expanded=false;
  this.children=new Array();
  this.level = function()
  {
    var node=this;
    var i=0;
    while(node.parentNode!=null){
      i++;
      node=node.parentNode;
    };
    return i;
  };
  this.hasChildren = function()
  {
    return this.children.length>0;
  };
  this.init = function()
  {
    var s = "";
    if(this.treeView.ns4){
      s = '<layer id="'+this.id()+'b" z-index="1" visibility="hidden">'+this.getContent()+'</layer>';
    }else{
      s = '<div id="'+this.id()+'b" style="position:absolute;visibility:hidden;z-index:1;">'+this.getContent()+'</div>';
    }
    return s;
  };
  this.getH = function()
  {
    return this.treeView.ns4?this.el.clip.height:this.el.offsetHeight;
  };
  this.getW = function()
  {
    return this.treeView.ns4?this.el.clip.width:this.el.offsetWidth;
  };
  this.id = function()
  {
    return 'nt'+this.treeView.name+this.index;
  };
  this.getContent = function()
  {
    var s = '';
    s+='<table border="0" cellpadding="0" cellspacing="0"><tr>';
    if(this.leCase=="BACK"){
      s+='<td background="'+this.treeView.fmt.bgG+'"><img src="'+this.treeView.fmt.nB+'" border="0" width="'+this.treeView.fmt.bgW+'" height="'+this.treeView.fmt.Bh+'"></td>';
      s+='<td background="'+this.treeView.fmt.bgF+'"><img src="'+this.treeView.fmt.nB+'" border="0" width="'+((laTailleMax+(this.treeView.fmt.left/2))-(2*this.treeView.fmt.bgW))+'" height="'+this.treeView.fmt.Bh+'"></td>';
      s+='<td background="'+this.treeView.fmt.bgD+'"><img src="'+this.treeView.fmt.nB+'" border="0" width="'+this.treeView.fmt.bgW+'" height="'+this.treeView.fmt.Bh+'"></td>';
    }
    if(this.leCase=="TOP"){
      s+='<td background="'+this.treeView.fmt.topG+'"><img src="'+this.treeView.fmt.nB+'" border="0" width="'+this.treeView.fmt.topW+'" height="'+this.treeView.fmt.topH+'"></td>';
      s+='<td background="'+this.treeView.fmt.topF+'"><img src="'+this.treeView.fmt.nB+'" border="0" width="'+((laTailleMax+(this.treeView.fmt.left/2))-(2*this.treeView.fmt.bgW))+'" height="'+this.treeView.fmt.topH+'"></td>';
      s+='<td background="'+this.treeView.fmt.topD+'"><img src="'+this.treeView.fmt.nB+'" border="0" width="'+this.treeView.fmt.topW+'" height="'+this.treeView.fmt.topH+'"></td>';
    }
    if(this.leCase=="BOT"){
      s+='<td background="'+this.treeView.fmt.botG+'"><img src="'+this.treeView.fmt.nB+'" border="0" width="'+this.treeView.fmt.botW+'" height="'+this.treeView.fmt.botH+'"></td>';
      s+='<td background="'+this.treeView.fmt.botF+'"><img src="'+this.treeView.fmt.nB+'" border="0" width="'+((laTailleMax+(this.treeView.fmt.left/2))-(2*this.treeView.fmt.bgW))+'" height="'+this.treeView.fmt.botH+'"></td>';
      s+='<td background="'+this.treeView.fmt.botD+'"><img src="'+this.treeView.fmt.nB+'" border="0" width="'+this.treeView.fmt.botW+'" height="'+this.treeView.fmt.botH+'"></td>';
    }
    s+='</tr></table>';
    return s;
  };
  this.moveTo = function( x, y )
  {
    if(this.treeView.ns4){
      this.el.moveTo(x,y);
    }else{
      this.el.style.left=x;
      this.el.style.top=y;
    }
  };
  this.show = function( sh )
  {
    if(this.visible==sh){
      return;
    }
    this.visible=sh;
    var vis=this.treeView.ns4?(sh?'show':'hide'):(sh?'visible':'hidden');
    if(this.treeView.ns4){
      this.el.visibility=vis;
    }else{
      this.el.style.visibility=vis;
    }
  };
  this.hideChildren = function()
  {
    this.show(false);
    for(var i=0; i<this.children.length; i++){
      this.children[i].hideChildren();
    };
  };
  this.draw = function()
  {
    var ll=this.treeView.fmt.left;
    this.moveTo(this.treeView.fmt.left, this.treeView.currTop);
    if(ll+this.getW()>this.treeView.maxWidth){
      this.treeView.maxWidth=ll+this.getW();
    }
    this.show(true);
    this.treeView.currTop+=this.getH();
    if(this.treeView.currTop>this.treeView.maxHeight){
      this.treeView.maxHeight=this.treeView.currTop;
    }
    if(this.expanded && this.hasChildren()){
      for(var i=0; i<this.children.length; i++){
        this.children[i].draw();
      };
    }
  };
};
//Prerequis
function und( val )
{
  return typeof(val)=='undefined';
};
window.oldCTOnLoad=window.onload;
window.onload=function ()
{
  var bw=new bw_check();
  if(bw.ns4 || bw.opera){
    window.origWidth=this.innerWidth;
    window.origHeight=this.innerHeight;
    if(bw.opera && und(window.operaRH)){
      window.operaRH=1;
      resizeHandler();
    }
  }
  if(window.oldCTOnLoad){
    window.oldCTOnLoad();
  }
};
function resizeHandler()
{
  var bw=new bw_check();
  if(this.innerWidth!=window.origWidth || this.innerHeight!=window.origHeight){
    location.reload();
  }
  if(bw.opera){
    setTimeout('resizeHandler()',500);
  }
};
if(new bw_check().ns4){
  window.onresize=resizeHandler;
}