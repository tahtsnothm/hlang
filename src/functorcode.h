






class FunctorExpr < Expr {
	public string name; // name created by line and counter maintained by the parser
	public void Expr::init (paramarray arr) {
		name = arr[0];
	} 
	public variant Expr::accept(ExprVisitor& visitor) {
		return visitor.visitFunctorExpr(me);
	}
	public variant Expr::copy() {
		FunctorExpr temp = new FunctorExpr;
		temp.Expr::init(name);
		return temp;
	}
}


func void callArgument_f() {
// try and parse an arrow function
	if (match(LEFT_SQUARE)) {
		FunctorExpr myfunctorExpr = new FunctorExpr;
		
		array inputs;
		arad(inputs, consume(IDENTIFIER, "Expected identifier"));
		while (match(COMMA)) {
			arad(inputs, consume(IDENTIFIER, "Expected identifier"));
		}
		consume(RIGHT_SQUARE, "Expected ']'");
		Token arrowkey = consume (ARROW_, "Expected arrow for arrow sign");
		
		string functorname = "functl" & previous().m_line & "c" & functorcounter; // takes from the arrow sign
		functorcounter++;
		myfunctorExpr.Expr::init (functorname);

		Stmt halfbody;
		if (match(LEFT_BRACE)) {
			halfbody = block_f();
		} else {
			Expr tempexpr = expression_();
			ReturnStmt tempbodyreturn = new ReturnStmt;
			tempbodyreturn.Stmt::init(arrowkey, tempexpr.copy(),true);
			halfbody = tempbodyreturn;
		}


		// generate the unroller
		array body;
		for (Token param : inputs ) (int i ) {
			AssignExpr tassignexpr = new AssignExpr;			
			VariableExpr inputname = new VariableExpr;
			AccessorExpr arraccessor = new AccessorExpr;
			VariableExpr arrname = new VariableExpr;
			Token arrnametoken = new Token;
			Token fakeBracket = new Token;
			LiteralExpr numberr = new LiteralExpr;
			
			ExprStmt wraps = new ExprStmt;


			arrnametoken.init(IDENTIFIER, "arr", "arr", arrowkey.m_line);
			fakeBracket.init(LEFT_SQUARE, "[fake", "[fake", arrowkey.m_line);
			numberr.Expr::init(i);
			arrname.Expr::init(arrnametoken);
			inputname.Expr::init(param);
			arraccessor.Expr::init(arrname, fakeBracket, array(numberr), "access", false );
			tassignexpr.Expr::init(inputname, arraccessor, false);			
			

			wraps.Stmt::init(tassignexpr);
			arad(body, wraps);
		}
		arad(body, halfbody);

		// generate the class
		ClassStmt functorClass = new ClassStmt;
		FunctionStmt fstmt = new FunctionStmt;
		fstmt.Stmt::init(mkident("funct_run", arrowkey.m_line), mkident("variant", arrowkey.m_line),
		 array(mkident("arr", arrowkey.m_line)), array (mkident("paramarray", arrowkey.m_line)), array(false), body, "public" );
		
		functorClass.Stmt::init(mkident(functorname, arrowkey.m_line), array(fstmt), array(mkident("funct", arrowkey.m_line)), array());

		
		//blabalbal
		arad(statements, functorClass);

		return myFunctorExpr;
	} else {
		return expression_();
	}
}

private func Token mkident (string name, long linenum) {
	Token t = new Token;

	t.init(IDENTIFIER, name, name, linenum);
	return t;
}



func variant ExprVisitor::visitFunctorExpr(FunctorExpr &expression) {
	write_forward_ ("dim m" & expression.name & " as new " & expression.name);
	
	return "m" & expression.name;
}

func variant map(variant &arr, funct &functor) {
	array newarr;
	for (variant item : arr) (int i) {
		arad(newarr, functor.run(item));
	}
	return newarr;
}

func variant filter(variant &arr, funct &functor) {
	array newarr;
	for (variant item : arr) (int i) {
		if (cbool(functor.run(item))) arad(newarr, item);
	}
	return newarr;
}
func variant reduce (variant &arr, funct &functor, variant base) {
	// [item, accum] => accum
	for (variant item : arr) (int i) {
		base = functor.run(item, base);
	}
	return base;
}

