var [%tree_name%]_FORMAT =
[
//0. left position
	10,
//1. top position
	110,
//2. show +/- buttons
	true,
//3. couple of button images (collapsed/expanded/blank)
	["[%tree_path%]tree_np.gif", "[%tree_path%]tree_nm.gif", "[%tree_path%]tree_nlp.gif", "[%tree_path%]tree_nlm.gif", "[%tree_path%]tree_nl.gif", "[%tree_path%]tree_nt.gif", "[%tree_path%]tree_nv.gif", "[%tree_path%]tree_blank.gif"],
//4. size of images (width, height,ident for nodes w/o children)
	[16,22,0],
//5. show folder image
	false,
//6. folder images (closed/opened/document)
	["[%tree_path%]tree_bf.gif", "[%tree_path%]tree_bo.gif", "[%tree_path%]tree_bd.gif"],
//7. size of images (width, height)
	[16,16],
//8. identation for each level [0/*first level*/, 16/*second*/, 32/*third*/,...]
	[[%tree_indent%]],
//9. tree background color ("" - transparent)
	"",
//10. default style for all nodes
	"clsDemoNode",
//11. styles for each level of menu (default style will be used for undefined levels)
	[],
//12. true if only one branch can be opened at same time
	false,
//13. item pagging and spacing
	[0,0],
//14. Active links on expand
	false,
//15. Active Background Image
	false,
//16. Background Images
	[],
//17. Background Images Sizes (width/height)
	[],
//18. Active Top Image
	false,
//19. Top Images
	[],
//20. Top Images Sizes (width/height)
	[],
//21. Active Bottom Image
	false,
//22. Bottom Images
	[],
//23. Bottom Images Sizes (width/height)
	[],
];
