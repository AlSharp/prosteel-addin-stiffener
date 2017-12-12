using System;
using System.Data;
using System.Collections.Generic;
using System.Text;

using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.Builders;
using MongoDB.Driver.Linq;

namespace GSFStiffener
{
	public static class StiffenerDb
	{
		public class Stiffener
		{
			public ObjectId Id {get; set;}
			public string Pos {get; set;}
			public float Height {get; set;}
			public float Width {get; set;}
			public float Chamfer {get; set;}
		}
		public static Stiffener GetStiffener(string positionNumber)
		{
			var connectionString = "mongodb://34.235.84.44:27017";
			var mongo_client = new MongoClient(connectionString);
			var server = mongo_client.GetServer();
			var database = server.GetDatabase("prosteel-db-development");
			MongoCollection<Stiffener> collection = database.GetCollection<Stiffener>("stiffeners");
			Stiffener stiff = new Stiffener();
			var query = Query.EQ("Pos", positionNumber);
			stiff = collection.FindOne(query);
			return stiff;
		}
	}
	public static class IBeamDb
	{
		public class IBeam
		{
			public int _id {get; set;}
			public string Type {get; set;}
			public bool Assigned {get; set;} 
		}
		public static IBeam GetIBeamByType(string type)
		{
			var connectionString = "mongodb://34.235.84.44:27017";
			var mongo_client = new MongoClient(connectionString);
			var server = mongo_client.GetServer();
			var database = server.GetDatabase("prosteel-db-development");
			MongoCollection<IBeam> collection = database.GetCollection<IBeam>("ibeams");
			IBeam beam = new IBeam();
			var query = Query.EQ("Type", type);
			beam = collection.FindOne(query);
			return beam;
		}
	}

	public static class StiffenerUtilities
	{
		public static void CreateOne()
		{

		}
	}

}