using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;

using dotnetconflib.Entity.Blog;
using dotnetconflib.Entity.Post;

namespace dotnetconflib.DBModel.dotnetconfDBContext
{
    public class dotnetconfContext : DbContext
    {
        public DbSet<Blog> Blogs { get; set; }
        public DbSet<Post> Posts { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(
                @"Server=servername, 1433;Database=2017dotnetconf;user=alanliu;password=password;");
        }
    }    
}