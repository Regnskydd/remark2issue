#!/usr/bin/python3

class Remark:
	def __init__(self, identifier, author, comment, decision, decision_comment, action_description, status):
		self.identifier = identifier
		self.author = author
		self.comment = comment
		self.decision = decision
		self.decision_comment = decision_comment
		self.action_description = action_description
		self.status = status
		
	def get_identifier(self):
		return self.identifier
		
	def get_author(self):
		return self.author
	
	def get_comment(self):
		return self.comment
	
	def get_decision(self):
		return self.decision
	
	def get_decision_comment(self):
		return self.decision_comment
		
	def get_action_description(self):
		return self.action_description
		
	def get_status(self):
		return self.status
