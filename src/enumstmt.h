string hello[3];
hello[0] = "s";
class EnumStmt < Stmt {
	public Token name;
	public variant types_;
	public void Stmt::init (variant arr) {
		name = arr[0];
		types_ = arr[1]; // types will be a list of tokens
	}

	public void Stmt::accept (StmtVistor visitor) {
		return visitor.visitEnumStmt(Me);
	}
	
}


func variant StmtVisitor::visitEnumStmt(EnumStmt statement) {
	write_("enum " & statement.name.m_lexeme);
	string s = "";
	for (Token item : statement.types_)(int i) {
		s = s & item.m_lexeme & " ";
	}
	if (len(s) > 2) {
		s = left(s, len(s) - 2);
	}
	indent();
	write_(s);
	unindent();
	write_("end enum");
}


func variant enumDeclaration_f() {
	Token name = consume(IDENTIFIER, "Expected enum name");
	array types_;	

	consume(LEFT_BRACE, "Expected '{' after enum");
	arad(types_, consume(IDENTIFIER, "enum needs at least one name"));
	while (match(COMMA)) {
		arad(types_, consume(IDENTIFIER, "Expected identifier"));
	}	
	consume(RIGHT_BRACE, "Expected '}' after enum declaration");

	EnumStmt temp = new EnumStmt;
	temp.stmt::init(name, types_);
	return temp;
}















persons = [person, person, ]
person = [bob, me7, lfdkjdf, kifdlj,]

persons.map({name, rank, age} => {

})



